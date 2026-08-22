"use client";

import React, { useRef, useState } from "react";
import {
  convertPngToPdf,
  convertPdfToPng,
  convertDocxToPdf,
  convertPdfToDocx,
  convertPngToJpg,
} from "@/utils/converter";

interface ToolItem {
  id: string;
  category: string;
  from: string;
  to: string;
  description: string;
  accept: string;
  multiple?: boolean;
}

const toolsData: ToolItem[] = [
  {
    id: "png-to-pdf",
    category: "Gambar ke Dokumen",
    from: "PNG",
    to: "PDF",
    description: "Gabungkan beberapa gambar PNG menjadi satu berkas PDF.",
    accept: "image/png",
    multiple: true,
  },
  {
    id: "pdf-to-png",
    category: "Dokumen ke Gambar",
    from: "PDF",
    to: "PNG",
    description: "Ubah setiap halaman PDF menjadi gambar PNG resolusi tinggi.",
    accept: "application/pdf",
  },
  {
    id: "docx-to-pdf",
    category: "Dokumen",
    from: "DOCX",
    to: "PDF",
    description: "Render dokumen Word menjadi PDF siap cetak ukuran A4.",
    accept: ".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  },
  {
    id: "pdf-to-docx",
    category: "Dokumen",
    from: "PDF",
    to: "DOCX",
    description: "Tarik teks dari PDF dan susun ulang menjadi dokumen Word.",
    accept: "application/pdf",
  },
  {
    id: "png-to-jpg",
    category: "Gambar",
    from: "PNG",
    to: "JPG",
    description: "Kompres PNG menjadi JPG dengan latar putih dan ukuran lebih ringan.",
    accept: "image/png",
    multiple: true,
  },
];

export default function Home() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [selectedTool, setSelectedTool] = useState<ToolItem>(toolsData[0]);
  const [status, setStatus] = useState<string>("");

  const handleCardClick = (tool: ToolItem) => {
    setSelectedTool(tool);
    setTimeout(() => {
      fileInputRef.current?.click();
    }, 50);
  };

  const downloadBlob = (blob: Blob, filename: string) => {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    a.click();
    URL.revokeObjectURL(url);
  };

  const getBaseName = (filename: string) => {
    return filename.substring(0, filename.lastIndexOf(".")) || filename;
  };

  const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    setStatus("Memproses konversi berkas...");
    try {
      const fileList = Array.from(files);
      if (selectedTool.id === "png-to-pdf") {
        const baseName = getBaseName(fileList[0].name);
        const pdfBlob = await convertPngToPdf(fileList);
        downloadBlob(pdfBlob, `${baseName}.pdf`);
      } else if (selectedTool.id === "pdf-to-png") {
        const baseName = getBaseName(fileList[0].name);
        const pngBlobs = await convertPdfToPng(fileList[0]);
        pngBlobs.forEach((blob, idx) => {
          downloadBlob(blob, `${baseName}-page-${idx + 1}.png`);
        });
      } else if (selectedTool.id === "docx-to-pdf") {
        const baseName = getBaseName(fileList[0].name);
        const pdfBlob = await convertDocxToPdf(fileList[0]);
        downloadBlob(pdfBlob, `${baseName}.pdf`);
      } else if (selectedTool.id === "pdf-to-docx") {
        const baseName = getBaseName(fileList[0].name);
        const docxBlob = await convertPdfToDocx(fileList[0]);
        downloadBlob(docxBlob, `${baseName}.docx`);
      } else if (selectedTool.id === "png-to-jpg") {
        const jpgBlobs = await convertPngToJpg(fileList);
        jpgBlobs.forEach((blob, idx) => {
          const baseName = getBaseName(fileList[idx].name);
          downloadBlob(blob, `${baseName}.jpg`);
        });
      }
      setStatus("Konversi selesai! Berkas berhasil diunduh.");
    } catch {
      setStatus("Gagal memproses konversi. Pastikan format berkas sesuai.");
    } finally {
      if (fileInputRef.current) {
        fileInputRef.current.value = "";
      }
    }
  };

  return (
    <div className="min-h-screen bg-[#F8FAFC] text-slate-900 font-sans pb-16 sm:pb-20">
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileChange}
        accept={selectedTool.accept}
        multiple={selectedTool.multiple}
        className="hidden"
      />

      <header className="flex items-center justify-between px-4 sm:px-8 py-4 sm:py-5 border-b border-slate-100 bg-white">
        <div className="flex items-center gap-2.5 sm:gap-3">
          <div className="w-7 h-7 sm:w-8 sm:h-8 flex items-center justify-center">
            <svg
              className="w-6 h-6 sm:w-7 sm:h-7 fill-slate-900"
              viewBox="0 0 10 8"
              shapeRendering="crispEdges"
            >
              <path d="M1 0h3v1H1zm5 0h3v1H6zM0 1h10v1H0zm0 1h10v1H0zm0 1h10v1H0zM1 4h8v1H1zM2 5h6v1H2zM3 6h4v1H3zM4 7h2v1H4z" />
            </svg>
          </div>
          <span className="font-extrabold text-lg sm:text-xl tracking-tight text-slate-900">
            iLoveSvvagy
          </span>
        </div>
      </header>

      <main className="max-w-4xl mx-auto px-4 sm:px-6 pt-10 sm:pt-16">
        <section className="text-center max-w-2xl mx-auto mb-10 sm:mb-16">
          <h1 className="text-3xl sm:text-5xl font-black text-slate-900 tracking-tight leading-[1.18] sm:leading-[1.15] mb-4 sm:mb-6">
            Convert berkas dengan presisi, langsung dari browser.
          </h1>
          <p className="text-slate-500 text-sm sm:text-lg leading-relaxed">
            Lima tool terpisah untuk lima kebutuhan. Pilih satu, taruh berkasnya, unduh hasilnya — tanpa akun dan tanpa mengunggah apa pun.
          </p>
          {status && (
            <div className="mt-5 sm:mt-6 p-3 bg-blue-50 border border-blue-200 text-blue-700 font-medium rounded-xl inline-block text-xs sm:text-sm">
              {status}
            </div>
          )}
        </section>

        <section className="space-y-4">
          <div className="mb-4 sm:mb-6">
            <h2 className="text-xl sm:text-2xl font-bold text-slate-900">Pilih Tool</h2>
            <p className="text-slate-500 text-xs sm:text-sm mt-1">
              Satu tombol untuk satu jenis konversi.
            </p>
          </div>

          <div className="grid gap-3 sm:gap-4">
            {toolsData.map((tool) => (
              <div
                key={tool.id}
                onClick={() => handleCardClick(tool)}
                className={`p-5 sm:p-6 rounded-2xl bg-white transition-all cursor-pointer border ${
                  selectedTool.id === tool.id
                    ? "border-blue-500 ring-1 ring-blue-500 shadow-sm"
                    : "border-slate-100 hover:border-slate-300 shadow-xs"
                }`}
              >
                <span className="text-[11px] sm:text-xs font-mono text-blue-600 bg-blue-50/50 px-2.5 py-1 rounded-md inline-block mb-2.5 sm:mb-3">
                  {tool.category}
                </span>
                <div className="flex items-center gap-2 mb-1.5 sm:mb-2">
                  <span className="text-lg sm:text-xl font-extrabold text-slate-900">
                    {tool.from}
                  </span>
                  <span className="text-blue-500 font-semibold">→</span>
                  <span className="text-lg sm:text-xl font-extrabold text-slate-900">
                    {tool.to}
                  </span>
                </div>
                <p className="text-slate-500 text-xs sm:text-sm">{tool.description}</p>
              </div>
            ))}
          </div>
        </section>
      </main>
    </div>
  );
}