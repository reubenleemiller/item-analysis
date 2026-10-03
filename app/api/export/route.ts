import { NextRequest, NextResponse } from "next/server";
import { createItemAnalysisWorkbook, type ExportPayload } from "@/lib/workbook";

export const runtime = "nodejs";

export async function POST(request: NextRequest) {
  try {
    const payload = (await request.json()) as ExportPayload;
    const workbook = await createItemAnalysisWorkbook(payload);
    const fileName = `${(payload.examTitle || "item-analysis").replace(/[^a-z0-9]+/gi, "-").replace(/(^-|-$)/g, "").toLowerCase() || "item-analysis"}.xlsx`;
    return new NextResponse(workbook, {
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${fileName}"`,
      },
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unable to create the workbook.";
    return NextResponse.json({ error: message }, { status: 400 });
  }
}
