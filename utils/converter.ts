import { jsPDF } from "jspdf";
import mammoth from "mammoth";
import { Document, Packer, Paragraph, TextRun, AlignmentType, PageOrientation } from "docx";

export const convertPngToPdf = async (files: File[]): Promise<Blob> => {
  const doc = new jsPDF();
  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const imageUrl = URL.createObjectURL(file);
    const img = new Image();
    img.src = imageUrl;
    await new Promise((resolve) => {
      img.onload = resolve;
    });
    if (i > 0) {
      doc.addPage();
    }
    const imgWidth = doc.internal.pageSize.getWidth();
    const imgHeight = (img.height * imgWidth) / img.width;
    doc.addImage(img, "PNG", 0, 0, imgWidth, imgHeight);
    URL.revokeObjectURL(imageUrl);
  }
  return doc.output("blob");
};

export const convertPdfToPng = async (file: File): Promise<Blob[]> => {
  const pdfjsLib = await import("pdfjs-dist");
  pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;

  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  const blobs: Blob[] = [];

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const viewport = page.getViewport({ scale: 2.0 });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d");
    canvas.height = viewport.height;
    canvas.width = viewport.width;

    if (context) {
      await page.render({ canvasContext: context, viewport, canvas }).promise;
      const blob = await new Promise<Blob | null>((resolve) =>
        canvas.toBlob((b) => resolve(b), "image/png")
      );
      if (blob) blobs.push(blob);
    }
  }
  return blobs;
};

export const convertDocxToPdf = async (file: File): Promise<Blob> => {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.convertToHtml({ arrayBuffer });
  const htmlContent = result.value;

  const container = document.createElement("div");
  container.style.position = "absolute";
  container.style.left = "-9999px";
  container.style.top = "-9999px";
  container.style.width = "794px";
  container.style.padding = "48px";
  container.style.background = "#ffffff";
  container.style.color = "#000000";
  container.style.fontFamily = "Arial, sans-serif";
  container.style.fontSize = "14px";
  container.style.lineHeight = "1.5";
  container.innerHTML = `
    <style>
      table { border-collapse: collapse; width: 100%; margin: 12px 0; }
      th, td { border: 1px solid #000; padding: 6px 10px; text-align: left; }
      p { margin: 8px 0; }
      h1, h2, h3, h4, h5, h6 { margin: 12px 0 6px 0; }
      img { max-width: 100%; height: auto; }
    </style>
    ${htmlContent}
  `;
  document.body.appendChild(container);

  const html2canvas = (await import("html2canvas")).default;
  const canvas = await html2canvas(container, {
    scale: 2,
    useCORS: true,
    logging: false,
  });

  document.body.removeChild(container);

  const imgData = canvas.toDataURL("image/png");
  const pdf = new jsPDF("p", "mm", "a4");
  const pdfWidth = pdf.internal.pageSize.getWidth();
  const pdfHeight = pdf.internal.pageSize.getHeight();
  const imgWidth = pdfWidth;
  const imgHeight = (canvas.height * pdfWidth) / canvas.width;

  let heightLeft = imgHeight;
  let position = 0;

  pdf.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight);
  heightLeft -= pdfHeight;

  while (heightLeft > 0) {
    position = heightLeft - imgHeight;
    pdf.addPage();
    pdf.addImage(imgData, "PNG", 0, position, imgWidth, imgHeight);
    heightLeft -= pdfHeight;
  }

  return pdf.output("blob");
};

export const convertPdfToDocx = async (file: File): Promise<Blob> => {
  const pdfjsLib = await import("pdfjs-dist");
  pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsLib.version}/build/pdf.worker.min.mjs`;

  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

  const paragraphs: Paragraph[] = [];

  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const viewport = page.getViewport({ scale: 1.0 });
    const textContent = await page.getTextContent();

    const linesMap = new Map<number, any[]>();

    for (const item of textContent.items as any[]) {
      if (!item.str || item.str.trim() === "") continue;
      const y = Math.round(item.transform[5]);
      let foundY = Array.from(linesMap.keys()).find((k) => Math.abs(k - y) <= 4);
      if (foundY === undefined) {
        foundY = y;
        linesMap.set(foundY, []);
      }
      linesMap.get(foundY)!.push(item);
    }

    const sortedYs = Array.from(linesMap.keys()).sort((a, b) => b - a);

    for (const y of sortedYs) {
      const items = linesMap.get(y)!;
      items.sort((a, b) => a.transform[4] - b.transform[4]);

      const lineText = items.map((it) => it.str).join(" ").replace(/\s+/g, " ").trim();
      if (!lineText) continue;

      const minX = items[0].transform[4];
      const maxX = items[items.length - 1].transform[4] + (items[items.length - 1].width || 0);
      const pageMiddle = viewport.width / 2;
      const lineMiddle = (minX + maxX) / 2;

      let alignment: (typeof AlignmentType)[keyof typeof AlignmentType] = AlignmentType.LEFT;
      if (Math.abs(lineMiddle - pageMiddle) < 40 && minX > 50) {
        alignment = AlignmentType.CENTER;
      } else if (minX > viewport.width * 0.6) {
        alignment = AlignmentType.RIGHT;
      }

      paragraphs.push(
        new Paragraph({
          alignment,
          spacing: { after: 120 },
          children: [
            new TextRun({
              text: lineText,
              font: "Arial",
              size: 22,
            }),
          ],
        })
      );
    }
  }

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: {
              width: 11906,
              height: 16838,
              orientation: PageOrientation.PORTRAIT,
            },
            margin: {
              top: 1440,
              right: 1440,
              bottom: 1440,
              left: 1440,
            },
          },
        },
        children: paragraphs,
      },
    ],
  });

  return await Packer.toBlob(doc);
};

export const convertPngToJpg = async (files: File[]): Promise<Blob[]> => {
  const blobs: Blob[] = [];
  for (const file of files) {
    const imageUrl = URL.createObjectURL(file);
    const img = new Image();
    img.src = imageUrl;
    await new Promise((resolve) => {
      img.onload = resolve;
    });
    const canvas = document.createElement("canvas");
    canvas.width = img.width;
    canvas.height = img.height;
    const ctx = canvas.getContext("2d");
    if (ctx) {
      ctx.fillStyle = "#FFFFFF";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0);
      const blob = await new Promise<Blob | null>((resolve) =>
        canvas.toBlob((b) => resolve(b), "image/jpeg", 0.92)
      );
      if (blob) blobs.push(blob);
    }
    URL.revokeObjectURL(imageUrl);
  }
  return blobs;
};