"use client";

import { useRef, useState } from "react";

type TargetResult = { name: string; rows: number; photos: number };

type Status =
  | { kind: "idle" }
  | { kind: "loading" }
  | { kind: "success"; results: TargetResult[] }
  | { kind: "error"; message: string };

const TARGETS = [
  { id: "priority", label: "Приоритет" },
  { id: "bitrix", label: "Битрикс" },
];

export default function Home() {
  const [status, setStatus] = useState<Status>({ kind: "idle" });
  const [fileName, setFileName] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  async function fetchAndDownload(
    file: File,
    target: string
  ): Promise<TargetResult> {
    const formData = new FormData();
    formData.append("file", file);

    const res = await fetch(`/api/transfer?target=${target}`, {
      method: "POST",
      body: formData,
    });

    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.error ?? `Ошибка сервера (${res.status})`);
    }

    const rows = Number(res.headers.get("X-Rows-Transferred") ?? 0);
    const photos = Number(res.headers.get("X-Photos-Transferred") ?? 0);
    const disposition = res.headers.get("Content-Disposition") ?? "";
    const nameMatch = disposition.match(/filename\*=UTF-8''([^;]+)/);
    const name = nameMatch
      ? decodeURIComponent(nameMatch[1])
      : `${target}.xlsx`;

    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    return { name, rows, photos };
  }

  async function handleFile(file: File) {
    setFileName(file.name);
    setStatus({ kind: "loading" });

    try {
      const results = await Promise.all(
        TARGETS.map((t) => fetchAndDownload(file, t.id))
      );
      setStatus({ kind: "success", results });
    } catch (err) {
      setStatus({
        kind: "error",
        message: err instanceof Error ? err.message : "Не удалось обработать файл.",
      });
    }
  }

  function onDrop(e: React.DragEvent<HTMLDivElement>) {
    e.preventDefault();
    const file = e.dataTransfer.files?.[0];
    if (file) handleFile(file);
  }

  return (
    <div className="flex min-h-screen flex-col items-center bg-zinc-50 px-6 py-16 font-sans dark:bg-black">
      <main className="flex w-full max-w-xl flex-col gap-8">
        <div className="text-center">
          <h1 className="text-2xl font-semibold tracking-tight text-zinc-900 dark:text-zinc-50">
            Перенос данных в спецификацию
          </h1>
          <p className="mt-2 text-sm leading-6 text-zinc-600 dark:text-zinc-400">
            Загрузите расчётную Excel-таблицу — данные и фото товаров
            автоматически перенесутся сразу в два готовых файла (Приоритет и
            Битрикс), которые скачаются один за другим.
          </p>
        </div>

        <div
          onDragOver={(e) => e.preventDefault()}
          onDrop={onDrop}
          onClick={() => inputRef.current?.click()}
          className="flex cursor-pointer flex-col items-center gap-3 rounded-2xl border-2 border-dashed border-zinc-300 bg-white px-6 py-14 text-center transition-colors hover:border-zinc-400 dark:border-zinc-700 dark:bg-zinc-900 dark:hover:border-zinc-600"
        >
          <input
            ref={inputRef}
            type="file"
            accept=".xlsx"
            className="hidden"
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) handleFile(file);
              e.target.value = "";
            }}
          />
          <span className="text-sm font-medium text-zinc-800 dark:text-zinc-200">
            Перетащите .xlsx файл сюда или нажмите, чтобы выбрать
          </span>
          <span className="text-xs text-zinc-500 dark:text-zinc-500">
            Поддерживается формат .xlsx
          </span>
        </div>

        {fileName && (
          <p className="text-center text-xs text-zinc-500 dark:text-zinc-500">
            Файл: {fileName}
          </p>
        )}

        {status.kind === "loading" && (
          <p className="text-center text-sm text-zinc-600 dark:text-zinc-400">
            Обработка файла…
          </p>
        )}

        {status.kind === "success" && (
          <div className="flex flex-col gap-2 rounded-xl border border-green-200 bg-green-50 px-4 py-3 text-center text-sm text-green-800 dark:border-green-900 dark:bg-green-950 dark:text-green-300">
            <span>Готово! Скачано {status.results.length} файла:</span>
            {status.results.map((r) => (
              <span key={r.name}>
                {r.name} — строк: {r.rows}, фото: {r.photos}
              </span>
            ))}
          </div>
        )}

        {status.kind === "error" && (
          <div className="rounded-xl border border-red-200 bg-red-50 px-4 py-3 text-center text-sm text-red-800 dark:border-red-900 dark:bg-red-950 dark:text-red-300">
            {status.message}
          </div>
        )}
      </main>
    </div>
  );
}
