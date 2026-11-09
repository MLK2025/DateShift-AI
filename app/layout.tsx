import type { Metadata } from "next";
import "./globals.css";

export const dynamic = "force-static";

const siteUrl =
  process.env.NEXT_PUBLIC_SITE_URL ??
  "https://dateshift-ai.handypixie6.chatgpt.site";

export const metadata: Metadata = {
  metadataBase: new URL(new URL(siteUrl).origin),
  title: "DateShift — Update dates and balance document figures",
  description: "Update dates and calculate exact target totals in DOCX and PDF documents directly in your browser.",
  openGraph: {
    title: "DateShift — Update dates. Balance figures. Keep the document.",
    description: "Update dates and rebalance financial figures to an exact total in Word and PDF documents.",
    images: [{ url: `${process.env.NEXT_PUBLIC_BASE_PATH ?? ""}/og.png`, width: 800, height: 420, alt: "DateShift AI" }],
  },
  twitter: {
    card: "summary_large_image",
    title: "DateShift — Update dates and balance document figures",
    description: "Update dates and rebalance financial figures to an exact total in Word and PDF documents.",
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
