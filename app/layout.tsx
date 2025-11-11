export const metadata = {
  title: "DOCX Formatter",
  description: "Normalize Word .docx formatting in one click"
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body style={{ fontFamily: 'system-ui, -apple-system, Segoe UI, Roboto, Ubuntu, Cantarell, Noto Sans, Arial, sans-serif', margin: 0 }}>
        {children}
      </body>
    </html>
  );
}
