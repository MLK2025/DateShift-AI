import type { Metadata } from "next";
import "./globals.css";

export const dynamic = "force-static";

const siteUrl =
  process.env.NEXT_PUBLIC_SITE_URL ??
  "https://dateshift-ai.handypixie6.chatgpt.site";

export const metadata: Metadata = {
  metadataBase: new URL(new URL(siteUrl).origin),
  title: "DateShift — Update document dates privately",
  description: "Find and update dates in DOCX and PDF documents directly in your browser, with no account or billing required.",
  openGraph: {
    title: "DateShift — Move every date. Keep every detail.",
    description: "Find and update dates in DOCX and PDF documents privately in your browser.",
    images: [{ url: `${process.env.NEXT_PUBLIC_BASE_PATH ?? ""}/og.png`, width: 800, height: 420, alt: "DateShift AI" }],
  },
  twitter: {
    card: "summary_large_image",
    title: "DateShift — Move every date. Keep every detail.",
    description: "Find and update dates in DOCX and PDF documents privately in your browser.",
    images: [`${process.env.NEXT_PUBLIC_BASE_PATH ?? ""}/og.png`],
  },
  icons: {
    icon: `${process.env.NEXT_PUBLIC_BASE_PATH ?? ""}/favicon.svg`,
    shortcut: `${process.env.NEXT_PUBLIC_BASE_PATH ?? ""}/favicon.svg`,
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
