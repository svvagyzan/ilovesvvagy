import { jsPDF } from "jspdf";
import mammoth from "mammoth";
import { Document, Packer, Paragraph, TextRun, AlignmentType, PageOrientation } from "docx";
import { PDFDocument } from "pdf-lib";
import JSZip from "jszip";

const uint8ToBase64 = (bytes: Uint8Array): string => {
  let binary = "";
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode.apply(null, Array.from(chunk));
  }
  return btoa(binary);
};

const extractImagesFromPptBytes = (bytes: Uint8Array): string[] => {
  const dataUrls: string[] = [];

  for (let i = 0; i < bytes.length - 8; i++) {
    if (
      bytes[i] === 0x89 &&
      bytes[i + 1] === 0x50 &&
      bytes[i + 2] === 0x4e &&
      bytes[i + 3] === 0x47 &&
      bytes[i + 4] === 0x0d &&
      bytes[i + 5] === 0x0a &&
      bytes[i + 6] === 0x1a &&
      bytes[i + 7] === 0x0a
    ) {
      let end = -1;
      for (let j = i + 8; j < bytes.length - 8; j++) {
        if (
          bytes[j] === 0x49 &&
          bytes[j + 1] === 0x45 &&
          bytes[j + 2] === 0x4e &&
          bytes[j + 3] === 0x44 &&
          bytes[j + 4] === 0xae &&
          bytes[j + 5] === 0x42 &&
          bytes[j + 6] === 0x60 &&
          bytes[j + 7] === 0x82
        ) {
          end = j + 8;
          break;
        }
      }
      if (end !== -1 && end - i > 500) {
        const slice = bytes.subarray(i, end);
        dataUrls.push(`data:image/png;base64,${uint8ToBase64(slice)}`);
        i = end;
      }
    }
  }

  for (let i = 0; i < bytes.length - 3; i++) {
    if (bytes[i] === 0xff && bytes[i + 1] === 0xd8 && bytes[i + 2] === 0xff) {
      let end = -1;
      for (let j = i + 3; j < bytes.length - 2; j++) {
        if (bytes[j] === 0xff && bytes[j + 1] === 0xd9) {
          end = j + 2;
          break;
        }
      }
      if (end !== -1 && end - i > 1000 && end - i < 10000000) {
        const slice = bytes.subarray(i, end);
        dataUrls.push(`data:image/jpeg;base64,${uint8ToBase64(slice)}`);
        i = end;
      }
    }
  }

  return dataUrls;
};

export const convertPngToPdf = async (files: File[]): Promise<Blob> => {
  let doc: jsPDF | null = null;

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const imageUrl = URL.createObjectURL(file);
    const img = new Image();
    img.src = imageUrl;
    await new Promise((resolve) => {
      img.onload = resolve;
    });

    const orientation = img.width > img.height ? "l" : "p";

    if (i === 0) {
      doc = new jsPDF({
        orientation,
        unit: "px",
        format: [img.width, img.height],
      });
    } else if (doc) {
      doc.addPage([img.width, img.height], orientation);
    }

    if (doc) {
      doc.addImage(img, "PNG", 0, 0, img.width, img.height);
    }

    URL.revokeObjectURL(imageUrl);
  }

  return doc ? doc.output("blob") : new Blob();
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
  container.style.boxSizing = "border-box";
  container.style.background = "#ffffff";
  container.style.color = "#000000";
  container.style.fontFamily = "Arial, sans-serif";
  container.style.fontSize = "14px";
  container.style.lineHeight = "1.5";
  container.innerHTML = htmlContent;
  document.body.appendChild(container);

  const images = Array.from(container.querySelectorAll("img"));
  await Promise.all(
    images.map(
      (img) =>
        new Promise((resolve) => {
          if (img.complete) {
            resolve(true);
          } else {
            img.onload = () => resolve(true);
            img.onerror = () => resolve(true);
          }
        })
    )
  );

  const PAGE_HEIGHT = 1123;
  const children = Array.from(container.children) as HTMLElement[];

  for (let i = 0; i < children.length; i++) {
    const child = children[i];
    const top = child.offsetTop;
    const height = child.offsetHeight;

    if (height === 0) continue;

    const startPage = Math.floor(top / PAGE_HEIGHT);
    const endPage = Math.floor((top + height - 1) / PAGE_HEIGHT);

    if (startPage !== endPage && top % PAGE_HEIGHT !== 0) {
      const nextPageTop = (startPage + 1) * PAGE_HEIGHT;
      const spacerHeight = nextPageTop - top;
      const spacer = document.createElement("div");
      spacer.style.height = `${spacerHeight}px`;
      spacer.style.width = "100%";
      container.insertBefore(spacer, child);
    }
  }

  const html2canvas = (await import("html2canvas")).default;
  const canvas = await html2canvas(container, {
    scale: 2,
    useCORS: true,
    logging: false,
  });

  document.body.removeChild(container);

  const pdf = new jsPDF("p", "mm", "a4");
  const pdfWidth = pdf.internal.pageSize.getWidth();
  const pdfHeight = pdf.internal.pageSize.getHeight();

  const totalHeightPx = canvas.height;
  const pageHeightPx = Math.round((canvas.width * 297) / 210);

  const pageCanvas = document.createElement("canvas");
  pageCanvas.width = canvas.width;
  pageCanvas.height = pageHeightPx;
  const pageCtx = pageCanvas.getContext("2d");

  let renderedHeight = 0;
  let pageIndex = 0;

  while (renderedHeight < totalHeightPx) {
    if (pageIndex > 0) {
      pdf.addPage();
    }

    if (pageCtx) {
      pageCtx.fillStyle = "#ffffff";
      pageCtx.fillRect(0, 0, pageCanvas.width, pageCanvas.height);
      pageCtx.drawImage(
        canvas,
        0,
        renderedHeight,
        canvas.width,
        pageHeightPx,
        0,
        0,
        canvas.width,
        pageHeightPx
      );
    }

    const pageImgData = pageCanvas.toDataURL("image/jpeg", 0.95);
    pdf.addImage(pageImgData, "JPEG", 0, 0, pdfWidth, pdfHeight);

    renderedHeight += pageHeightPx;
    pageIndex++;
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

export const convertPptxToPdf = async (file: File): Promise<Blob> => {
  const arrayBuffer = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(arrayBuffer);

  const mediaMap = new Map<string, string>();
  const mediaFiles: string[] = [];
  zip.forEach((relativePath) => {
    if (relativePath.startsWith("ppt/media/")) {
      mediaFiles.push(relativePath);
    }
  });

  for (const mediaPath of mediaFiles) {
    const base64Data = await zip.file(mediaPath)?.async("base64");
    if (base64Data) {
      const ext = mediaPath.split(".").pop()?.toLowerCase();
      const mime = ext === "png" ? "image/png" : "image/jpeg";
      const fileName = mediaPath.split("/").pop() || "";
      mediaMap.set(fileName, `data:${mime};base64,${base64Data}`);
    }
  }

  const slideFiles: string[] = [];
  zip.forEach((relativePath) => {
    if (relativePath.match(/^ppt\/slides\/slide\d+\.xml$/)) {
      slideFiles.push(relativePath);
    }
  });

  slideFiles.sort((a, b) => {
    const numA = parseInt(a.match(/\d+/)![0], 10);
    const numB = parseInt(b.match(/\d+/)![0], 10);
    return numA - numB;
  });

  const pdf = new jsPDF({
    orientation: "l",
    unit: "px",
    format: [1280, 720],
  });

  const html2canvas = (await import("html2canvas")).default;

  for (let i = 0; i < slideFiles.length; i++) {
    const slidePath = slideFiles[i];
    const slideXmlText = await zip.file(slidePath)?.async("string");
    if (!slideXmlText) continue;

    const relsPath = slidePath.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels";
    const relsXmlText = await zip.file(relsPath)?.async("string");

    const relsMap = new Map<string, string>();
    if (relsXmlText) {
      const parser = new DOMParser();
      const relsDoc = parser.parseFromString(relsXmlText, "text/xml");
      const relationships = Array.from(relsDoc.getElementsByTagName("Relationship"));
      relationships.forEach((rel) => {
        const id = rel.getAttribute("Id");
        const target = rel.getAttribute("Target");
        if (id && target) {
          const fileName = target.split("/").pop() || "";
          relsMap.set(id, fileName);
        }
      });
    }

    const parser = new DOMParser();
    const xmlDoc = parser.parseFromString(slideXmlText, "text/xml");

    const paragraphs: string[] = [];
    const pNodes = Array.from(xmlDoc.getElementsByTagName("a:p"));
    for (const pNode of pNodes) {
      const tNodes = Array.from(pNode.getElementsByTagName("a:t"));
      const lineText = tNodes.map((n) => n.textContent || "").join(" ").trim();
      if (lineText.length > 0) {
        paragraphs.push(lineText);
      }
    }

    const slideImages: string[] = [];
    const blipNodes = Array.from(xmlDoc.getElementsByTagName("a:blip"));
    for (const blip of blipNodes) {
      const embedId = blip.getAttribute("r:embed");
      if (embedId && relsMap.has(embedId)) {
        const fileName = relsMap.get(embedId)!;
        if (mediaMap.has(fileName)) {
          slideImages.push(mediaMap.get(fileName)!);
        }
      }
    }

    const tableRows: string[][] = [];
    const trNodes = Array.from(xmlDoc.getElementsByTagName("a:tr"));
    for (const trNode of trNodes) {
      const rowData: string[] = [];
      const tcNodes = Array.from(trNode.getElementsByTagName("a:tc"));
      for (const tcNode of tcNodes) {
        const tNodes = Array.from(tcNode.getElementsByTagName("a:t"));
        const cellText = tNodes.map((n) => n.textContent || "").join(" ").trim();
        rowData.push(cellText);
      }
      if (rowData.length > 0) {
        tableRows.push(rowData);
      }
    }

    const container = document.createElement("div");
    container.style.position = "absolute";
    container.style.left = "-9999px";
    container.style.top = "-9999px";
    container.style.width = "1280px";
    container.style.height = "720px";
    container.style.background = "#ffffff";
    container.style.color = "#1e293b";
    container.style.padding = "40px 60px";
    container.style.boxSizing = "border-box";
    container.style.fontFamily = "Arial, sans-serif";
    container.style.display = "flex";
    container.style.flexDirection = "column";
    container.style.justifyContent = "flex-start";
    container.style.alignItems = "stretch";
    container.style.border = "1px solid #e2e8f0";

    if (paragraphs.length > 0) {
      const titleText = paragraphs[0];
      const banner = document.createElement("div");
      banner.style.backgroundColor = "#a3e635";
      banner.style.padding = "10px 24px";
      banner.style.borderRadius = "4px";
      banner.style.marginBottom = "24px";
      banner.style.alignSelf = "flex-start";

      const h1 = document.createElement("h1");
      h1.style.fontSize = "32px";
      h1.style.fontWeight = "bold";
      h1.style.color = "#000000";
      h1.style.margin = "0";
      h1.innerText = titleText;
      banner.appendChild(h1);
      container.appendChild(banner);
    }

    const contentBox = document.createElement("div");
    contentBox.style.flex = "1";
    contentBox.style.display = "flex";
    contentBox.style.flexDirection = "column";
    contentBox.style.justifyContent = "center";

    if (slideImages.length > 0) {
      for (const imgUri of slideImages) {
        const imgEl = document.createElement("img");
        imgEl.src = imgUri;
        imgEl.style.maxWidth = "900px";
        imgEl.style.maxHeight = "480px";
        imgEl.style.width = "auto";
        imgEl.style.height = "auto";
        imgEl.style.objectFit = "contain";
        imgEl.style.margin = "0 auto";
        imgEl.style.display = "block";
        imgEl.style.borderRadius = "8px";
        contentBox.appendChild(imgEl);
      }
    }

    if (tableRows.length > 0) {
      const table = document.createElement("table");
      table.style.width = "100%";
      table.style.borderCollapse = "collapse";
      table.style.fontSize = "18px";
      table.style.marginTop = "16px";

      tableRows.forEach((row, rIdx) => {
        const tr = document.createElement("tr");
        if (rIdx === 0) {
          tr.style.background = "#94a3b8";
          tr.style.color = "#ffffff";
          tr.style.fontWeight = "bold";
        } else {
          tr.style.background = rIdx % 2 === 0 ? "#f8fafc" : "#ffffff";
        }

        row.forEach((cellText) => {
          const cell = document.createElement(rIdx === 0 ? "th" : "td");
          cell.style.padding = "14px";
          cell.style.border = "1px solid #cbd5e1";
          cell.style.textAlign = "left";
          cell.innerText = cellText;
          tr.appendChild(cell);
        });
        table.appendChild(tr);
      });
      contentBox.appendChild(table);
    }

    if (paragraphs.length > 1) {
      for (let pIdx = 1; pIdx < paragraphs.length; pIdx++) {
        const p = document.createElement("p");
        p.style.fontSize = "18px";
        p.style.lineHeight = "1.6";
        p.style.color = "#1e293b";
        p.style.marginBottom = "16px";
        p.innerText = paragraphs[pIdx];
        contentBox.appendChild(p);
      }
    }

    container.appendChild(contentBox);
    document.body.appendChild(container);

    const canvas = await html2canvas(container, {
      scale: 2,
      useCORS: true,
      logging: false,
    });

    document.body.removeChild(container);

    const imgData = canvas.toDataURL("image/jpeg", 0.95);
    if (i > 0) {
      pdf.addPage([1280, 720], "l");
    }
    pdf.addImage(imgData, "JPEG", 0, 0, 1280, 720);
  }

  return pdf.output("blob");
};

export const convertPptToPdf = async (file: File): Promise<Blob> => {
  const arrayBuffer = await file.arrayBuffer();
  const bytes = new Uint8Array(arrayBuffer);

  const extractedImages = extractImagesFromPptBytes(bytes);

  const cleanTexts: string[] = [];
  for (let i = 0; i < bytes.length - 8; i++) {
    const recType = bytes[i + 2] | (bytes[i + 3] << 8);
    const recLen =
      bytes[i + 4] |
      (bytes[i + 5] << 8) |
      (bytes[i + 6] << 16) |
      (bytes[i + 7] << 24);

    if (recLen > 0 && recLen < 50000 && i + 8 + recLen <= bytes.length) {
      if (recType === 0x0fa8) {
        let text = "";
        for (let j = 0; j < recLen; j++) {
          const charCode = bytes[i + 8 + j];
          if (charCode >= 32 && charCode <= 126) {
            text += String.fromCharCode(charCode);
          } else if (charCode === 10 || charCode === 13) {
            text += "\n";
          }
        }
        const clean = text.trim();
        if (clean.length > 1) {
          cleanTexts.push(clean);
        }
      } else if (recType === 0x0fa0) {
        let text = "";
        for (let j = 0; j < recLen; j += 2) {
          const charCode = bytes[i + 8 + j] | (bytes[i + 8 + j + 1] << 8);
          if (
            (charCode >= 32 && charCode <= 0xd7ff) ||
            charCode === 10 ||
            charCode === 13
          ) {
            text += String.fromCharCode(charCode);
          }
        }
        const clean = text.trim();
        if (clean.length > 1) {
          cleanTexts.push(clean);
        }
      }
    }
  }

  interface DynamicSlide {
    paragraphs: string[];
    imageUri?: string;
  }

  const slides: DynamicSlide[] = [];
  const chunkSize = 3;

  for (let i = 0; i < cleanTexts.length; i += chunkSize) {
    const slideParagraphs = cleanTexts.slice(i, i + chunkSize);
    let imageUri: string | undefined = undefined;
    if (extractedImages.length > slides.length) {
      imageUri = extractedImages[slides.length];
    }
    slides.push({
      paragraphs: slideParagraphs,
      imageUri,
    });
  }

  if (slides.length === 0) {
    if (extractedImages.length > 0) {
      extractedImages.forEach((imgUri) => {
        slides.push({
          paragraphs: [],
          imageUri: imgUri,
        });
      });
    } else {
      slides.push({
        paragraphs: ["Dokumen PPT berhasil dikonversi."],
      });
    }
  }

  const pdf = new jsPDF({
    orientation: "l",
    unit: "px",
    format: [1280, 720],
  });

  const html2canvas = (await import("html2canvas")).default;

  for (let i = 0; i < slides.length; i++) {
    const slide = slides[i];

    const container = document.createElement("div");
    container.style.position = "absolute";
    container.style.left = "-9999px";
    container.style.top = "-9999px";
    container.style.width = "1280px";
    container.style.height = "720px";
    container.style.background = "#ffffff";
    container.style.color = "#1e293b";
    container.style.padding = "40px 60px";
    container.style.boxSizing = "border-box";
    container.style.fontFamily = "Arial, sans-serif";
    container.style.display = "flex";
    container.style.flexDirection = "column";
    container.style.justifyContent = "flex-start";
    container.style.alignItems = "stretch";
    container.style.border = "1px solid #e2e8f0";

    if (slide.paragraphs.length > 0) {
      const banner = document.createElement("div");
      banner.style.backgroundColor = "#a3e635";
      banner.style.padding = "10px 24px";
      banner.style.borderRadius = "4px";
      banner.style.marginBottom = "24px";
      banner.style.alignSelf = "flex-start";

      const h1 = document.createElement("h1");
      h1.style.fontSize = "32px";
      h1.style.fontWeight = "bold";
      h1.style.color = "#000000";
      h1.style.margin = "0";
      h1.innerText = slide.paragraphs[0];
      banner.appendChild(h1);
      container.appendChild(banner);
    }

    const contentBox = document.createElement("div");
    contentBox.style.flex = "1";
    contentBox.style.display = "flex";
    contentBox.style.flexDirection = "column";
    contentBox.style.justifyContent = "center";

    if (slide.imageUri) {
      const imgEl = document.createElement("img");
      imgEl.src = slide.imageUri;
      imgEl.style.maxWidth = "900px";
      imgEl.style.maxHeight = "480px";
      imgEl.style.width = "auto";
      imgEl.style.height = "auto";
      imgEl.style.objectFit = "contain";
      imgEl.style.margin = "0 auto";
      imgEl.style.display = "block";
      imgEl.style.borderRadius = "8px";
      contentBox.appendChild(imgEl);
    }

    if (slide.paragraphs.length > 1) {
      for (let pIdx = 1; pIdx < slide.paragraphs.length; pIdx++) {
        const p = document.createElement("p");
        p.style.fontSize = "18px";
        p.style.lineHeight = "1.6";
        p.style.color = "#1e293b";
        p.style.marginBottom = "16px";
        p.innerText = slide.paragraphs[pIdx];
        contentBox.appendChild(p);
      }
    }

    container.appendChild(contentBox);
    document.body.appendChild(container);

    const canvas = await html2canvas(container, {
      scale: 2,
      useCORS: true,
      logging: false,
    });

    document.body.removeChild(container);

    const imgData = canvas.toDataURL("image/jpeg", 0.95);
    if (i > 0) {
      pdf.addPage([1280, 720], "l");
    }
    pdf.addImage(imgData, "JPEG", 0, 0, 1280, 720);
  }

  return pdf.output("blob");
};

export const mergePdfs = async (files: File[]): Promise<Blob> => {
  const mergedPdf = await PDFDocument.create();

  for (const file of files) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await PDFDocument.load(arrayBuffer);
    const copiedPages = await mergedPdf.copyPages(pdf, pdf.getPageIndices());
    copiedPages.forEach((page) => mergedPdf.addPage(page));
  }

  const pdfBytes = await mergedPdf.save();
  return new Blob([pdfBytes as unknown as BlobPart], {
    type: "application/pdf",
  });
};