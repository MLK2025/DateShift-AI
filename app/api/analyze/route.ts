import { NextResponse } from "next/server";

export const dynamic = "force-dynamic";

type DateShiftRuntime = typeof globalThis & {
  __DATESHIFT_OPENAI_API_KEY?: string;
  __DATESHIFT_OPENAI_MODEL?: string;
};

const MAX_FILE_BYTES = 25 * 1024 * 1024;
const SUPPORTED_EXTENSIONS = new Set(["doc", "docx", "pdf"]);

const SYSTEM_PROMPT = `You are the document-date engine for DateShift AI.
Read the entire supplied document, including visible text in scanned PDF pages. Identify calendar dates and follow the user's editing instruction exactly.

Rules:
- Return only dates that the instruction says to update. Do not treat invoice numbers, reference numbers, amounts, page numbers, versions, or standalone years as editable dates unless the instruction explicitly asks for them.
- Preserve each date's original written format unless the user asks for another format.
- Recognize weekday-prefixed dates such as Wed Mar 04 2026, and recompute the weekday when the date changes.
- Recognize timestamp dates such as Wed Mar 04 15:57:15 GMT 2026. Preserve the time and timezone while shifting the calendar date.
- When shifting by months, preserve the day when possible. If that day does not exist in the destination month, use the destination month's final day.
- Use nearby document context to interpret ambiguous numeric dates. Never invent a date that is not present.
- Context must be a short quote or faithful summary that helps the user locate the date.
- Explain exclusions or ambiguity briefly in assistant_message.
- If no dates match the instruction, return an empty dates array and explain why.
- Produce only the required structured response.`;

const outputSchema = {
  type: "object",
  additionalProperties: false,
  properties: {
    assistant_message: { type: "string" },
    interpreted_instruction: { type: "string" },
    dates: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        properties: {
          original: { type: "string" },
          updated: { type: "string" },
          context: { type: "string" },
          reason: { type: "string" },
          confidence: { type: "number", minimum: 0, maximum: 100 },
        },
        required: ["original", "updated", "context", "reason", "confidence"],
      },
    },
  },
  required: ["assistant_message", "interpreted_instruction", "dates"],
} as const;

function fileExtension(name: string) {
  return name.split(".").pop()?.toLowerCase() ?? "";
}

function mimeTypeFor(file: File) {
  if (file.type) return file.type;
  const extension = fileExtension(file.name);
  if (extension === "pdf") return "application/pdf";
  if (extension === "docx") return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
  return "application/msword";
}

function bytesToBase64(bytes: Uint8Array) {
  const chunkSize = 0x8000;
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return btoa(binary);
}

function responseText(payload: Record<string, unknown>) {
  if (typeof payload.output_text === "string") return payload.output_text;
  const output = Array.isArray(payload.output) ? payload.output : [];
  for (const item of output) {
    if (!item || typeof item !== "object") continue;
    const content = Array.isArray((item as { content?: unknown[] }).content)
      ? (item as { content: unknown[] }).content
      : [];
    for (const part of content) {
      if (part && typeof part === "object" && typeof (part as { text?: unknown }).text === "string") {
        return (part as { text: string }).text;
      }
    }
  }
  return "";
}

export async function POST(request: Request) {
  try {
    const runtime = globalThis as DateShiftRuntime;
    const apiKey = process.env.OPENAI_API_KEY ?? runtime.__DATESHIFT_OPENAI_API_KEY;
    if (!apiKey) {
      return NextResponse.json(
        { error: "The secure AI connection is not configured yet. Add the OpenAI API key to enable document analysis." },
        { status: 503 },
      );
    }

    const form = await request.formData();
    const suppliedFile = form.get("file");
    const file = suppliedFile instanceof File && suppliedFile.size ? suppliedFile : null;
    const demoText = String(form.get("demo_text") ?? "").trim();
    const instruction = String(form.get("instruction") ?? "").trim();
    const months = Math.max(1, Math.min(24, Number(form.get("months") ?? 3) || 3));
    const direction = form.get("direction") === "backward" ? "backward" : "forward";
    const effectiveInstruction = instruction || `Move every calendar date ${direction} by ${months} month${months === 1 ? "" : "s"}.`;

    if (!file && !demoText) {
      return NextResponse.json({ error: "Choose a Word or PDF document before asking the AI to analyze it." }, { status: 400 });
    }

    const userContent: Array<Record<string, unknown>> = [];
    if (file) {
      const extension = fileExtension(file.name);
      if (!SUPPORTED_EXTENSIONS.has(extension)) {
        return NextResponse.json({ error: "Only Word (.doc or .docx) and PDF documents are supported." }, { status: 400 });
      }
      if (file.size > MAX_FILE_BYTES) {
        return NextResponse.json({ error: "This document is larger than the 25 MB upload limit." }, { status: 413 });
      }
      const bytes = new Uint8Array(await file.arrayBuffer());
      userContent.push({
        type: "input_file",
        filename: file.name,
        file_data: `data:${mimeTypeFor(file)};base64,${bytesToBase64(bytes)}`,
      });
    } else {
      userContent.push({ type: "input_text", text: `Sample document contents:\n\n${demoText}` });
    }

    userContent.push({
      type: "input_text",
      text: `User instruction:\n${effectiveInstruction}\n\nThe monthly controls currently show ${months} month${months === 1 ? "" : "s"} ${direction}. Use those controls only as a fallback when the written instruction does not provide a different rule.`,
    });

    const openAIResponse = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: process.env.OPENAI_MODEL ?? runtime.__DATESHIFT_OPENAI_MODEL ?? "gpt-6-astra",
        store: false,
        input: [
          { role: "system", content: [{ type: "input_text", text: SYSTEM_PROMPT }] },
          { role: "user", content: userContent },
        ],
        text: {
          format: {
            type: "json_schema",
            name: "document_date_updates",
            strict: true,
            schema: outputSchema,
          },
        },
      }),
    });

    const payload = await openAIResponse.json() as Record<string, unknown>;
    if (!openAIResponse.ok) {
      const apiError = payload.error && typeof payload.error === "object"
        ? String((payload.error as { message?: unknown }).message ?? "OpenAI could not analyze this document.")
        : "OpenAI could not analyze this document.";
      return NextResponse.json({ error: apiError }, { status: openAIResponse.status });
    }

    const rawText = responseText(payload);
    if (!rawText) {
      return NextResponse.json({ error: "The AI completed the request but returned no readable result. Please try again." }, { status: 502 });
    }

    const parsed = JSON.parse(rawText) as {
      assistant_message: string;
      interpreted_instruction: string;
      dates: Array<{ original: string; updated: string; context: string; reason: string; confidence: number }>;
    };

    return NextResponse.json(parsed);
  } catch (error) {
    const message = error instanceof Error ? error.message : "The document could not be analyzed.";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
