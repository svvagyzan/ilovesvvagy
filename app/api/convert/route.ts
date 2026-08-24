import { NextResponse } from "next/server";
import {
  PDFServices,
  ServicePrincipalCredentials,
  CreatePDFJob,
  CreatePDFResult,
  MimeType,
} from "@adobe/pdfservices-node-sdk";
import fs from "fs";
import path from "path";
import os from "os";

export async function POST(req: Request) {
  let tempInputPath = "";

  try {
    const formData = await req.formData();
    const file = formData.get("file") as File;

    if (!file) {
      return NextResponse.json(
        { error: "File tidak ditemukan" },
        { status: 400 }
      );
    }

    const credentials = new ServicePrincipalCredentials({
      clientId: process.env.ADOBE_CLIENT_ID || "",
      clientSecret: process.env.ADOBE_CLIENT_SECRET || "",
    });

    const pdfServices = new PDFServices({ credentials });

    const bytes = await file.arrayBuffer();
    const buffer = Buffer.from(bytes);

    const uniqueId = `${Date.now()}_${Math.random().toString(36).substring(2, 7)}`;
    tempInputPath = path.join(os.tmpdir(), `input_${uniqueId}_${file.name}`);
    fs.writeFileSync(tempInputPath, buffer);

    const readStream = fs.createReadStream(tempInputPath);
    const inputAsset = await pdfServices.upload({
      readStream: readStream as any,
      mimeType: MimeType.DOCX,
    });

    const job = new CreatePDFJob({ inputAsset });
    const pollingURL = await pdfServices.submit({ job });

    const pdfServicesResponse = await pdfServices.getJobResult({
      pollingURL,
      resultType: CreatePDFResult,
    });

    const resultAsset = pdfServicesResponse.result?.asset;
    if (!resultAsset) {
      throw new Error("Gagal mendapatkan hasil konversi dari Adobe API");
    }

    const streamAsset = await pdfServices.getContent({ asset: resultAsset });

    const chunks: Uint8Array[] = [];
    for await (const chunk of streamAsset.readStream as any) {
      chunks.push(typeof chunk === "string" ? Buffer.from(chunk) : chunk);
    }
    const pdfBuffer = Buffer.concat(chunks);

    if (fs.existsSync(tempInputPath)) {
      fs.unlinkSync(tempInputPath);
    }

    return new Response(pdfBuffer, {
      headers: {
        "Content-Type": "application/pdf",
        "Content-Disposition": `attachment; filename="${file.name.replace(/\.docx$/i, "")}.pdf"`,
      },
    });
  } catch (error: any) {
    if (tempInputPath && fs.existsSync(tempInputPath)) {
      fs.unlinkSync(tempInputPath);
    }

    return NextResponse.json(
      { error: error.message || "Gagal mengkonversi dokumen" },
      { status: 500 }
    );
  }
}