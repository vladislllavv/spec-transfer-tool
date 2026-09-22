import { NextRequest, NextResponse } from "next/server";
import { transferSpecification } from "@/lib/transfer";

export const runtime = "nodejs";
export const maxDuration = 60;

export async function POST(req: NextRequest) {
  const formData = await req.formData();
  const file = formData.get("file");

  if (!file || !(file instanceof File)) {
    return NextResponse.json(
      { error: "Файл не найден. Загрузите Excel-файл (.xlsx)." },
      { status: 400 }
    );
  }

  const arrayBuffer = await file.arrayBuffer();
  const buffer = Buffer.from(arrayBuffer);

  try {
    const result = await transferSpecification(buffer);
    return new NextResponse(new Uint8Array(result.buffer), {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition":
          "attachment; filename=\"Specification.xlsx\"; filename*=UTF-8''" +
          encodeURIComponent("Спецификация_Готовая.xlsx"),
        "X-Rows-Transferred": String(result.rowsTransferred),
        "X-Photos-Transferred": String(result.photosTransferred),
      },
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : "Неизвестная ошибка";
    return NextResponse.json({ error: message }, { status: 422 });
  }
}
