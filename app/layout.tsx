import type { Metadata } from "next";
import "./styles.css";

export const metadata: Metadata = {
  title: "RM Tutoring Services | Item Analysis Builder",
  description: "Create an Item Analysis Google Sheet from an exam setup and class list.",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
