import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "SheetFlow CRM Starter",
  description: "A professional website starter connected to Google Sheets through a secure backend.",
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

