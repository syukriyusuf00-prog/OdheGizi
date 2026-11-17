/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { jsPDF } from "jspdf";
import { SPPGProfile } from "../types";
import { formatIndonesianDate } from "../utils/menuSchedule";
import { formatRupiah } from "../utils/calc";

export type ArchiveCategory = "jumlah_pm" | "menu" | "food_cost" | "nota_pesanan";

export interface ArchivePMRecord {
  totalPM: number;
  totalKecil: number;
  totalBesar: number;
  totalAlergi: number;
  rows: Array<{
    label: string;
    porsiKecil: number;
    porsiBesar: number;
    alergiKecil: number;
    alergiBesar: number;
    total: number;
  }>;
}

export interface ArchiveMenuPlanRecord {
  planTitle: string;
  dietType: string;
  sasaranLabel: string;
  namaMenu: string;
  karbohidrat: string;
  laukHewani: string;
  laukNabati: string;
  sayur: string;
  buah: string;
  susu: string;
}

export interface ArchiveFoodCostSection {
  jenisMenu: string;
  namaMenu: string;
  bufferPct: number;
  pmKecil: number;
  pmBesar: number;
  rabTotal: number;
  actualTotal: number;
  costPerPorsiKecil: number;
  costPerPorsiBesar: number;
  energiBesar: number;
  proteinBesar: number;
  energiKecil: number;
  proteinKecil: number;
  bahanBesar: Array<{
    nama: string;
    beratBB: number;
    urt: string;
    kebutuhanKg: number;
    jumlahBuffer: number;
    hargaSatuan: number;
    hargaTotal: number;
  }>;
  bahanKecil: Array<{
    nama: string;
    beratBB: number;
    urt: string;
    kebutuhanKg: number;
    jumlahBuffer: number;
    hargaSatuan: number;
    hargaTotal: number;
  }>;
}

export interface ArchiveNotaItemRecord {
  nama: string;
  jumlah: number;
  satuan: string;
  hargaSatuan: number;
  total: number;
}

export interface DailyArchiveEntry {
  id: string;
  monthKey: string; // YYYY-MM
  dateStr: string; // YYYY-MM-DD
  hariKe: number;
  primaryMenuName: string;
  sourceLabel: string; // "Perencanaan Aktif" or saved period name
  pm: ArchivePMRecord;
  menus: ArchiveMenuPlanRecord[];
  foodCosts: ArchiveFoodCostSection[];
  notaItems: ArchiveNotaItemRecord[];
  notaTotalBelanja: number;
}

export interface ArchivePdfOptions {
  entry: DailyArchiveEntry;
  category: ArchiveCategory;
  profile: SPPGProfile;
  kopLines: [string, string, string, string];
}

const CATEGORY_TITLES: Record<ArchiveCategory, string> = {
  jumlah_pm: "ARSIP JUMLAH PENERIMA MANFAAT (PM)",
  menu: "ARSIP PERENCANAAN MASTER MENU HARIAN",
  food_cost: "ARSIP KALKULASI FOOD COST & KECUKUPAN GIZI",
  nota_pesanan: "ARSIP NOTA PESANAN LOGISTIK BAHAN MAKANAN",
};

const CATEGORY_FILENAME_PREFIX: Record<ArchiveCategory, string> = {
  jumlah_pm: "Jumlah_PM",
  menu: "Master_Menu",
  food_cost: "Food_Cost",
  nota_pesanan: "Nota_Pesanan",
};

export function buildArchivePdfFilename(entry: DailyArchiveEntry, category: ArchiveCategory): string {
  const safeMenu = (entry.primaryMenuName || `Menu_Hari_${entry.hariKe}`)
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, "_")
    .slice(0, 48);
  const safeDate = entry.dateStr || `Hari_${entry.hariKe}`;
  return `${CATEGORY_FILENAME_PREFIX[category]}_${safeDate}_${safeMenu}.pdf`;
}

function drawKopHeader(
  pdf: jsPDF,
  kopLines: [string, string, string, string],
  profile: SPPGProfile,
  subtitle: string
): number {
  const pageWidth = pdf.internal.pageSize.getWidth();
  const centerX = pageWidth / 2;
  let y = 14;

  pdf.setTextColor(15, 23, 42);
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(12);
  pdf.text((kopLines[0] || "BADAN GIZI NASIONAL").toUpperCase(), centerX, y, { align: "center" });
  y += 5;

  pdf.setFontSize(10.5);
  pdf.text((kopLines[1] || "SATUAN PELAYANAN PEMENUHAN GIZI").toUpperCase(), centerX, y, { align: "center" });
  y += 5;

  pdf.setFontSize(11);
  pdf.text((kopLines[2] || profile.namaLembaga || "SPPG").toUpperCase(), centerX, y, { align: "center" });
  y += 4.5;

  pdf.setFont("helvetica", "italic");
  pdf.setFontSize(8.5);
  pdf.setTextColor(71, 85, 105);
  const line4 = kopLines[3] || `Alamat: ${profile.alamat || "-"}`;
  const wrappedKop4 = pdf.splitTextToSize(line4, pageWidth - 28);
  pdf.text(wrappedKop4, centerX, y, { align: "center" });
  y += wrappedKop4.length * 3.8 + 2;

  pdf.setDrawColor(15, 23, 42);
  pdf.setLineWidth(0.7);
  pdf.line(14, y, pageWidth - 14, y);
  pdf.setLineWidth(0.2);
  pdf.line(14, y + 1.2, pageWidth - 14, y + 1.2);
  y += 7;

  // Document Banner
  pdf.setFillColor(17, 36, 97);
  pdf.roundedRect(14, y, pageWidth - 28, 8.5, 1.5, 1.5, "F");
  pdf.setTextColor(255, 255, 255);
  pdf.setFont("helvetica", "bold");
  pdf.setFontSize(10);
  pdf.text(subtitle, centerX, y + 5.7, { align: "center" });
  y += 13;

  return y;
}

function drawMetaBox(pdf: jsPDF, entry: DailyArchiveEntry, profile: SPPGProfile, startY: number): number {
  const pageWidth = pdf.internal.pageSize.getWidth();
  const boxWidth = pageWidth - 28;
  const formattedDate = formatIndonesianDate(entry.dateStr) || entry.dateStr;

  pdf.setFillColor(248, 250, 252);
  pdf.setDrawColor(203, 213, 225);
  pdf.roundedRect(14, startY, boxWidth, 18, 1.5, 1.5, "FD");

  pdf.setTextColor(30, 41, 59);
  pdf.setFontSize(8.5);

  pdf.setFont("helvetica", "bold");
  pdf.text("Tanggal Perencanaan :", 18, startY + 5.5);
  pdf.setFont("helvetica", "normal");
  pdf.text(`${formattedDate} (Hari Ke-${entry.hariKe})`, 56, startY + 5.5);

  pdf.setFont("helvetica", "bold");
  pdf.text("Nama Menu Utama :", 18, startY + 11);
  pdf.setFont("helvetica", "normal");
  const menuText = pdf.splitTextToSize(entry.primaryMenuName || `Menu Hari Ke-${entry.hariKe}`, boxWidth - 46);
  pdf.text(menuText[0] || "-", 56, startY + 11);

  pdf.setFont("helvetica", "bold");
  pdf.text("SPPG / Sumber Arsip :", 18, startY + 15.8);
  pdf.setFont("helvetica", "normal");
  pdf.text(`${profile.namaLembaga || "-"}  •  ${entry.sourceLabel}`, 56, startY + 15.8);

  return startY + 23;
}

function ensurePageSpace(
  pdf: jsPDF,
  currentY: number,
  neededHeight: number,
  kopLines: [string, string, string, string],
  profile: SPPGProfile,
  subtitle: string
): number {
  const pageHeight = pdf.internal.pageSize.getHeight();
  if (currentY + neededHeight <= pageHeight - 16) {
    return currentY;
  }
  pdf.addPage();
  return drawKopHeader(pdf, kopLines, profile, subtitle);
}

function drawSimpleTable(
  pdf: jsPDF,
  startY: number,
  columns: Array<{ label: string; width: number; align?: "left" | "center" | "right" }>,
  rows: Array<string[]>,
  kopLines: [string, string, string, string],
  profile: SPPGProfile,
  subtitle: string
): number {
  let y = startY;
  const startX = 14;
  const totalWidth = columns.reduce((acc, c) => acc + c.width, 0);

  const drawHeaderRow = (curY: number) => {
    pdf.setFillColor(226, 232, 240);
    pdf.setDrawColor(148, 163, 184);
    pdf.rect(startX, curY, totalWidth, 7, "FD");
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(7.5);
    pdf.setTextColor(15, 23, 42);

    let x = startX;
    columns.forEach((col) => {
      if (col.align === "right") {
        pdf.text(col.label, x + col.width - 2, curY + 4.7, { align: "right" });
      } else if (col.align === "center") {
        pdf.text(col.label, x + col.width / 2, curY + 4.7, { align: "center" });
      } else {
        pdf.text(col.label, x + 2, curY + 4.7);
      }
      pdf.line(x, curY, x, curY + 7);
      x += col.width;
    });
    pdf.line(startX + totalWidth, curY, startX + totalWidth, curY + 7);
    return curY + 7;
  };

  y = ensurePageSpace(pdf, y, 16, kopLines, profile, subtitle);
  y = drawHeaderRow(y);

  pdf.setFont("helvetica", "normal");
  pdf.setFontSize(7.5);

  rows.forEach((row, rIdx) => {
    const cellLines = row.map((cell, cIdx) =>
      pdf.splitTextToSize(String(cell ?? "-"), Math.max(columns[cIdx].width - 3.5, 6))
    );
    const maxLines = Math.max(1, ...cellLines.map((l) => l.length));
    const rowHeight = Math.max(6, maxLines * 3.5 + 2.5);

    if (y + rowHeight > pdf.internal.pageSize.getHeight() - 16) {
      pdf.addPage();
      y = drawKopHeader(pdf, kopLines, profile, subtitle);
      y = drawHeaderRow(y);
      pdf.setFont("helvetica", "normal");
      pdf.setFontSize(7.5);
    }

    if (rIdx % 2 === 1) {
      pdf.setFillColor(248, 250, 252);
      pdf.rect(startX, y, totalWidth, rowHeight, "F");
    }
    pdf.setDrawColor(203, 213, 225);
    pdf.rect(startX, y, totalWidth, rowHeight, "S");

    let x = startX;
    pdf.setTextColor(30, 41, 59);
    columns.forEach((col, cIdx) => {
      const lines = cellLines[cIdx];
      lines.forEach((ln: string, lineIdx: number) => {
        const textY = y + 4 + lineIdx * 3.5;
        if (col.align === "right") {
          pdf.text(ln, x + col.width - 2, textY, { align: "right" });
        } else if (col.align === "center") {
          pdf.text(ln, x + col.width / 2, textY, { align: "center" });
        } else {
          pdf.text(ln, x + 2, textY);
        }
      });
      pdf.line(x, y, x, y + rowHeight);
      x += col.width;
    });
    pdf.line(startX + totalWidth, y, startX + totalWidth, y + rowHeight);
    y += rowHeight;
  });

  return y + 4;
}

function drawSignatureFooter(pdf: jsPDF, profile: SPPGProfile, startY: number): void {
  const pageWidth = pdf.internal.pageSize.getWidth();
  const pageHeight = pdf.internal.pageSize.getHeight();
  let y = startY + 6;
  if (y + 28 > pageHeight - 10) {
    pdf.addPage();
    y = 22;
  }

  pdf.setTextColor(30, 41, 59);
  pdf.setFontSize(8.5);
  pdf.setFont("helvetica", "normal");
  pdf.text("Mengetahui,", 26, y);
  pdf.text("Disusun Oleh,", pageWidth - 68, y);

  pdf.setFont("helvetica", "bold");
  pdf.text("Kepala Satuan Pelayanan (SPPG)", 26, y + 4.5);
  pdf.text("Ahli Gizi / Nutritionist", pageWidth - 68, y + 4.5);

  pdf.text(profile.namaKepala || "(....................................)", 26, y + 20);
  pdf.text(profile.namaAhliGizi || "(....................................)", pageWidth - 68, y + 20);

  pdf.setFont("helvetica", "italic");
  pdf.setFontSize(7);
  pdf.setTextColor(100, 116, 139);
  pdf.text(
    `Diarsipkan secara otomatis oleh SISPERMBG pada ${new Date().toLocaleString("id-ID")}`,
    14,
    pageHeight - 8
  );
}

export function generateArchivePdfBlob(options: ArchivePdfOptions): Blob {
  const { entry, category, profile, kopLines } = options;
  const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const title = CATEGORY_TITLES[category];

  let y = drawKopHeader(pdf, kopLines, profile, title);
  y = drawMetaBox(pdf, entry, profile, y);

  if (category === "jumlah_pm") {
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(9.5);
    pdf.setTextColor(15, 23, 42);
    pdf.text(
      `Ringkasan Penerima Manfaat: Total ${entry.pm.totalPM} PM (Porsi Kecil: ${entry.pm.totalKecil} | Porsi Besar: ${entry.pm.totalBesar} | Alergi: ${entry.pm.totalAlergi})`,
      14,
      y
    );
    y += 5;

    const rows =
      entry.pm.rows.length > 0
        ? entry.pm.rows.map((r, idx) => [
            String(idx + 1),
            r.label,
            String(r.porsiKecil),
            String(r.porsiBesar),
            String(r.alergiKecil + r.alergiBesar),
            String(r.total),
          ])
        : [["-", "Belum ada data kelompok sasaran", "0", "0", "0", "0"]];

    y = drawSimpleTable(
      pdf,
      y,
      [
        { label: "No", width: 10, align: "center" },
        { label: "Kelompok Sasaran Penerima Manfaat", width: 88, align: "left" },
        { label: "Porsi Kecil", width: 22, align: "center" },
        { label: "Porsi Besar", width: 22, align: "center" },
        { label: "Alergi", width: 18, align: "center" },
        { label: "Total PM", width: 22, align: "center" },
      ],
      rows,
      kopLines,
      profile,
      title
    );
  } else if (category === "menu") {
    const rows =
      entry.menus.length > 0
        ? entry.menus.map((m, idx) => [
            String(idx + 1),
            `${m.planTitle} (${m.dietType})\n${m.sasaranLabel}`,
            m.namaMenu || "-",
            m.karbohidrat || "-",
            m.laukHewani || "-",
            m.laukNabati || "-",
            m.sayur || "-",
            [m.buah, m.susu].filter(Boolean).join(" / ") || "-",
          ])
        : [["-", "-", "Belum ada susunan menu pada tanggal ini", "-", "-", "-", "-", "-"]];

    y = drawSimpleTable(
      pdf,
      y,
      [
        { label: "No", width: 8, align: "center" },
        { label: "Kategori / Kelompok", width: 34, align: "left" },
        { label: "Nama Menu", width: 32, align: "left" },
        { label: "Pokok", width: 20, align: "left" },
        { label: "L. Hewani", width: 22, align: "left" },
        { label: "L. Nabati", width: 22, align: "left" },
        { label: "Sayur", width: 22, align: "left" },
        { label: "Buah / Susu", width: 22, align: "left" },
      ],
      rows,
      kopLines,
      profile,
      title
    );
  } else if (category === "food_cost") {
    if (entry.foodCosts.length === 0) {
      pdf.setFont("helvetica", "italic");
      pdf.setFontSize(9);
      pdf.text("Belum ada data Food Cost pada tanggal perencanaan ini.", 14, y + 4);
      y += 12;
    } else {
      entry.foodCosts.forEach((fc) => {
        y = ensurePageSpace(pdf, y, 24, kopLines, profile, title);
        pdf.setFillColor(238, 242, 255);
        pdf.roundedRect(14, y, pdf.internal.pageSize.getWidth() - 28, 12, 1.5, 1.5, "F");
        pdf.setFont("helvetica", "bold");
        pdf.setFontSize(8.5);
        pdf.setTextColor(30, 58, 138);
        pdf.text(
          `Kategori: ${fc.jenisMenu}  |  Menu: ${fc.namaMenu || entry.primaryMenuName}  |  Buffer: ${fc.bufferPct}%`,
          17,
          y + 4.8
        );
        pdf.setFont("helvetica", "normal");
        pdf.setFontSize(8);
        pdf.setTextColor(51, 65, 85);
        pdf.text(
          `PM Besar: ${fc.pmBesar} (${formatRupiah(fc.costPerPorsiBesar)}/porsi)  |  PM Kecil: ${fc.pmKecil} (${formatRupiah(fc.costPerPorsiKecil)}/porsi)  |  RAB: ${formatRupiah(fc.rabTotal)}  |  Realisasi: ${formatRupiah(fc.actualTotal)}`,
          17,
          y + 9.5
        );
        y += 15;

        const renderPorsiSection = (
          label: string,
          items: ArchiveFoodCostSection["bahanBesar"],
          energi: number,
          protein: number
        ) => {
          y = ensurePageSpace(pdf, y, 14, kopLines, profile, title);
          pdf.setFont("helvetica", "bold");
          pdf.setFontSize(8.5);
          pdf.setTextColor(15, 23, 42);
          pdf.text(
            `${label} (Energi: ${energi.toFixed(0)} kkal | Protein: ${protein.toFixed(1)} g)`,
            14,
            y
          );
          y += 3;

          const rows =
            items.length > 0
              ? items.map((b, i) => [
                  String(i + 1),
                  b.nama || "-",
                  `${b.beratBB} g`,
                  b.urt || "-",
                  `${b.kebutuhanKg.toFixed(2)} Kg`,
                  `${b.jumlahBuffer.toFixed(1)}`,
                  formatRupiah(b.hargaSatuan),
                  formatRupiah(b.hargaTotal),
                ])
              : [["-", "Tidak ada bahan", "-", "-", "-", "-", "-", "-"]];

          y = drawSimpleTable(
            pdf,
            y,
            [
              { label: "No", width: 9, align: "center" },
              { label: "Nama Bahan", width: 52, align: "left" },
              { label: "BB", width: 16, align: "center" },
              { label: "URT", width: 22, align: "center" },
              { label: "Keb. Kg", width: 20, align: "right" },
              { label: "Jml+Buf", width: 18, align: "right" },
              { label: "Harga", width: 22, align: "right" },
              { label: "Total", width: 23, align: "right" },
            ],
            rows,
            kopLines,
            profile,
            title
          );
        };

        if (fc.bahanBesar.length > 0 || fc.pmBesar > 0) {
          renderPorsiSection("Porsi Besar", fc.bahanBesar, fc.energiBesar, fc.proteinBesar);
        }
        if (fc.bahanKecil.length > 0 || fc.pmKecil > 0) {
          renderPorsiSection("Porsi Kecil", fc.bahanKecil, fc.energiKecil, fc.proteinKecil);
        }
      });
    }
  } else if (category === "nota_pesanan") {
    const rows =
      entry.notaItems.length > 0
        ? entry.notaItems.map((n, idx) => [
            String(idx + 1),
            n.nama,
            String(n.jumlah),
            n.satuan || "Kg",
            formatRupiah(n.hargaSatuan),
            formatRupiah(n.total),
          ])
        : [["-", "Belum ada item pesanan bahan makanan pada tanggal ini", "-", "-", "-", "-"]];

    y = drawSimpleTable(
      pdf,
      y,
      [
        { label: "No", width: 12, align: "center" },
        { label: "Uraian Jenis Bahan Makanan", width: 76, align: "left" },
        { label: "Jumlah", width: 20, align: "center" },
        { label: "Satuan", width: 20, align: "center" },
        { label: "Harga Satuan", width: 26, align: "right" },
        { label: "Total Harga", width: 28, align: "right" },
      ],
      rows,
      kopLines,
      profile,
      title
    );

    y = ensurePageSpace(pdf, y, 12, kopLines, profile, title);
    pdf.setFillColor(236, 253, 245);
    pdf.setDrawColor(16, 185, 129);
    pdf.roundedRect(14, y, pdf.internal.pageSize.getWidth() - 28, 9, 1.5, 1.5, "FD");
    pdf.setFont("helvetica", "bold");
    pdf.setFontSize(9.5);
    pdf.setTextColor(6, 95, 70);
    pdf.text("TOTAL BELANJA NOTA PESANAN:", 18, y + 6);
    pdf.text(formatRupiah(entry.notaTotalBelanja), pdf.internal.pageSize.getWidth() - 18, y + 6, {
      align: "right",
    });
    y += 12;
  }

  drawSignatureFooter(pdf, profile, y);
  return pdf.output("blob");
}
