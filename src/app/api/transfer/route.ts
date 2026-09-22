import { NextRequest, NextResponse } from "next/server";
import {
  transferToTemplate,
  TEMPLATE_PROFILES,
  type TemplateProfile,
} from "@/lib/transfer";

export const runtime = "nodejs";
export const maxDuration = 60;

function resolveProfile(req: NextRequest): TemplateProfile | null {
  const target = req.nextUrl.searchParams.get("target") ?? "priority";
  return TEMPLATE_PROFILES.find((p) => p.id === target) ?? null;
}

export async function POST(req: NextRequest) {
  const profile = resolveProfile(req);
  if (!profile) {
    return NextResponse.json(
      { error: "Неизвестный целевой шаблон (target)." },
      { status: 400 }
    );
  }

  let formData: FormData;
  try {
    formData = await req.formData();
  } catch {
    return NextResponse.json(
      { error: "Ожидается multipart/form-data с полем file." },
      { status: 400 }
    );
  }

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
    const result = await transferToTemplate(buffer, profile);
    const filename = `${result.outputName}.xlsx`;
    return new NextResponse(new Uint8Array(result.buffer), {
      status: 200,
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition":
          "attachment; filename=\"Specification.xlsx\"; filename*=UTF-8''" +
          encodeURIComponent(filename),
        "X-Rows-Transferred": String(result.rowsTransferred),
        "X-Photos-Transferred": String(result.photosTransferred),
      },
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : "Неизвестная ошибка";
    return NextResponse.json({ error: message }, { status: 422 });
  }
}
