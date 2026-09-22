"use client";

import { useRef, useState } from "react";

type Status =
  | { kind: "idle" }
  | { kind: "loading" }
  | { kind: "success"; rows: number; photos: number }
  | { kind: "error"; message: string };

export default function Home() {
  const [status, setStatus] = useState<Status>({ kind: "idle" });
  const [fileName, setFileName] = useState<string | null>(null);
  const inputRef = useRef<HTMLInputElement>(null);

  async function handleFile(file: File) {
    setFileName(file.name);
    setStatus({ kind: "loading" });

    const formData = new FormData();
    formData.append("file", file);

    try {
      const res = await fetch("/api/transfer", {
        method: "POST",
        body: formData,
      });

      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        setStatus({
          kind: "error",
          message: data.error ?? `Ошибка сервера (${res.status})`,
        });
        return;
      }

      const rows = Number(res.headers.get("X-Rows-Transferred") ?? 0);
      const photos = Number(res.headers.get("X-Photos-Transferred") ?? 0);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "Спецификация_Готовая.xlsx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      setStatus({ kind: "success", rows, photos });
    } catch {
      setStatus({ kind: "error", message: "Не удалось обработать файл." });
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
            автоматически перенесутся в шаблон спецификации, и готовый файл
            сразу скачается.
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
          <div className="rounded-xl border border-green-200 bg-green-50 px-4 py-3 text-center text-sm text-green-800 dark:border-green-900 dark:bg-green-950 dark:text-green-300">
            Готово! Перенесено строк: {status.rows}, фотографий:{" "}
            {status.photos}. Файл скачан.
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
