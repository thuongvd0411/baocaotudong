
import React, { useState, useEffect, useRef } from 'react';
import mammoth from 'mammoth';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import * as XLSX from 'xlsx';
import { GoogleGenAI, Type, GenerateContentResponse } from "@google/genai";
import { Document, Packer, Paragraph, Table, TableCell, TableRow, WidthType, TextRun, AlignmentType, VerticalAlign, BorderStyle } from "docx";
import { ESDMResult, ProcessingStatus, StudentInfo } from './types';
import Button from './components/Button';
import StatusAlert from './components/StatusAlert';
import ProgressBar from './components/ProgressBar';

// --- SECURITY CONFIG ---
const API_LICENSE_URL = "https://script.google.com/macros/s/AKfycbzojyLK8je1IsaOZWh18ljiw4Nb7sQt4wcWITrn6HmRIAAw2iZ0sw0Z4RBWqf3JIdeDwA/exec";
const APP_ID = "ESDM_EXPERT_PRO_V2";
const BUILD_ID = "2024_SECURE_BUILD";

// --- SECURITY HELPER FUNCTIONS ---
const getDeviceId = () => {
  let deviceId = localStorage.getItem('deviceId');
  if (!deviceId) {
    deviceId = crypto.randomUUID();
    localStorage.setItem('deviceId', deviceId);
  }
  return deviceId;
};

const generateAppFingerprint = async () => {
  const components = [
    APP_ID,
    BUILD_ID,
    navigator.userAgent,
    navigator.platform,
    `${window.screen.width}x${window.screen.height}`,
    // Use length of main_controller as a code integrity check
    (typeof main_controller === 'function' ? main_controller.toString().length : 0).toString()
  ];
  const msg = components.join('||');
  const msgBuffer = new TextEncoder().encode(msg);
  const hashBuffer = await crypto.subtle.digest('SHA-256', msgBuffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
};

// --- TYPE DEFINITIONS COMMON ---

type CoreAction = 'CALCULATE_AGE' | 'ANALYZE' | 'GENERATE_DOCX' | 'PARSE_EXCEL' | 'GENERATE_IEP' | 'GENERATE_REPORT' | 'MODULE_4_ANALYZE' | 'MODULE_4_FIX';

// --- TYPES FOR MODULE 2 ---
export type GoalSuffix = '(MTNT)' | '(MTC)' | '(MTP)';
export interface SelectedGoal { id: string; suffix: GoalSuffix; }
export interface Selection { level: string; domain: string; goals: SelectedGoal[]; }
export interface EsdmGoal { id: string; text: string; }
export interface EsdmDomain { name: string; goals: EsdmGoal[]; }
export interface EsdmLevel { name: string; domains: EsdmDomain[]; }
export interface ProcessedGoal {
  domainName: string; levelName: string; goalId: string; longTermGoal: string; suffix: string;
}

// --- TYPES FOR MODULE 3 ---
export interface Mod3Goal {
  id: string; // for UI key
  goal: string;
  percentage: number;
  note: string;
}
export interface Mod3FieldGroup {
  id: string; // for UI key
  fieldName: string;
  goals: Mod3Goal[];
}
export interface Mod3ChildInfo {
  name: string;
  dob: string;
  reportMonth: string;
  caregiverTitle: 'bố' | 'mẹ' | 'bố mẹ';
}
export interface Module3Data {
  childInfo: Mod3ChildInfo;
  fieldGroups: Mod3FieldGroup[];
}

// --- TYPES FOR MODULE 4 ---
export interface Mod4TableInfo {
  id: number;
  index: number; // Index in XML
  previewHtml: string; // From Mammoth
  issues: string[]; // Detected issues
  canMergeNext: boolean; // Is followed by another table closely?
  isMergeTarget: boolean; // Is selected to be merged into previous?
  options: {
    fixBorders: boolean;
    fixSpacing: boolean;
    autofit: boolean;
    mergeNext: boolean;
    fixAlign: boolean;
  };
}

interface CoreInput {
  action: CoreAction;
  payload: {
    // For CALCULATE_AGE
    dob?: string;
    evalDate?: string;
    ageFormat?: 'detail' | 'month';
    
    // For ANALYZE (Module 1)
    files?: File[];
    selectedLevels?: number[];
    selectedColumns?: number[];
    apiKey?: string;
    onProgress?: (percent: number) => void;

    // For GENERATE_DOCX (Module 1)
    templateFile?: File;
    studentInfo?: StudentInfo;
    esdmResult?: ESDMResult;
    fixCounter?: number;

    // For PARSE_EXCEL (Module 2)
    file?: File;

    // For GENERATE_IEP (Module 2)
    originalFileName?: string;
    selections?: Selection[];
    esdmData?: EsdmLevel[];
    smartSplitting?: boolean;

    // For GENERATE_REPORT (Module 3)
    module3Data?: Module3Data;

    // For MODULE_4
    mod4File?: File;
    mod4TableConfig?: Mod4TableInfo[];
  };
}

interface CoreOutput {
  age?: string;
  esdmResult?: ESDMResult;
  blob?: Blob;
  filename?: string;
  
  // Output for Module 2
  levelsData?: EsdmLevel[];
  
  // Output for Module 3 (reusing blob/filename)

  // Output for Module 4
  mod4Tables?: Mod4TableInfo[];
}

// --- GLOBAL HELPERS ---
function escapeXml(unsafe: string): string {
  return unsafe.replace(/[<>&'"]/g, (c) => {
    switch (c) { case '<': return '&lt;'; case '>': return '&gt;'; case '&': return '&amp;'; case '\'': return '&apos;'; case '"': return '&quot;'; default: return c; }
  });
}

function removeVietnameseTones(str: string): string {
  str = str.toLowerCase();
  str = str.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, "a");
  str = str.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, "e");
  str = str.replace(/ì|í|ị|ỉ|ĩ/g, "i");
  str = str.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, "o");
  str = str.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, "u");
  str = str.replace(/ỳ|ý|ỵ|ỷ|ỹ/g, "y");
  str = str.replace(/đ/g, "d");
  str = str.replace(/\u0300|\u0301|\u0303|\u0309|\u0323/g, ""); 
  str = str.replace(/\u02C6|\u0306|\u031B/g, ""); 
  return str;
}

function isGoalRowIndicator(val: any): boolean {
  const str = String(val || '').trim();
  if (!str) return false;
  return /^(M|m)?\d+(\.\d+)*\.?$/.test(str);
}

function isTableHeader(str: string): boolean {
  const lower = str.toLowerCase();
  const stopWords = ['stt', 'no.', 'mục tiêu smart', 'mục tiêu', 'nội dung', 'lĩnh vực', 'mã', 'code', 'mô tả', 'ghi chú', 'nhận xét', 'kết quả', 'đạt', 'chưa đạt', 'ngày'];
  return stopWords.some(w => lower === w || lower.startsWith(w + ' '));
}

function isDomainLabel(val: any): boolean {
  if (val === undefined || val === null) return false;
  const str = String(val).trim();
  if (str === '') return false;
  if (isGoalRowIndicator(str)) return false;
  if (isTableHeader(str)) return false;
  if (!isNaN(Number(str)) && !str.includes(' ')) return false; 
  return true;
}

// ============================================================================
// MODULE 4: THE MATRIX 2.0 & REPORT (INDEPENDENT MODULE)
// ============================================================================
async function autism_module_4(input: CoreInput): Promise<CoreOutput> {
  const { action, payload } = input;

  if (action === 'MODULE_4_ANALYZE') {
    const { mod4File } = payload;
    if (!mod4File) throw new Error("Chưa chọn file.");

    // 1. Get HTML Preview via Mammoth
    const arrayBuffer = await mod4File.arrayBuffer();
    const mammothResult = await mammoth.convertToHtml({ arrayBuffer });
    const fullHtml = mammothResult.value;
    
    // Split HTML to find tables. Note: This is an approximation for preview.
    // We create a temporary DOM to extract tables.
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = fullHtml;
    const htmlTables = Array.from(tempDiv.getElementsByTagName('table'));

    // 2. Parse XML to detect structure issues
    const pzip = new PizZip(arrayBuffer);
    const xmlStr = pzip.file("word/document.xml")?.asText() || "";
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlStr, "application/xml");
    
    const xmlTables = Array.from(doc.getElementsByTagName("w:tbl"));
    
    // 3. Map and Detect Issues
    const tableInfos: Mod4TableInfo[] = xmlTables.map((tbl, idx) => {
      const issues: string[] = [];
      const tblPr = tbl.getElementsByTagName("w:tblPr")[0];
      
      // Check Borders
      const borders = tblPr?.getElementsByTagName("w:tblBorders")[0];
      if (!borders) {
        issues.push("Thiếu viền bảng");
      } else {
        const sides = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'];
        const missing = sides.some(s => !borders.getElementsByTagName(`w:${s}`).length);
        if (missing) issues.push("Viền không đầy đủ");
      }

      // Check if mergeable
      let canMergeNext = false;
      if (idx < xmlTables.length - 1) {
        let sibling = tbl.nextSibling;
        let gapCount = 0;
        let isCleanGap = true;
        
        while (sibling && sibling !== xmlTables[idx+1]) {
           if (sibling.nodeName === 'w:p') {
             const text = sibling.textContent || "";
             if (text.trim().length > 0) { isCleanGap = false; break; }
             gapCount++;
           } else if (sibling.nodeName === 'w:tbl') {
             break; 
           }
           sibling = sibling.nextSibling;
        }
        
        if (isCleanGap && gapCount < 5) {
           canMergeNext = true;
           issues.push("Có thể gộp với bảng dưới");
        }
      }

      return {
        id: Date.now() + idx,
        index: idx,
        previewHtml: htmlTables[idx] ? htmlTables[idx].outerHTML : "<p class='text-xs text-gray-400'>Không thể tạo bản xem trước</p>",
        issues,
        canMergeNext,
        isMergeTarget: false, 
        // DO NOT AUTO-TICK OPTIONS
        options: {
          fixBorders: false,
          fixSpacing: false,
          autofit: false,
          mergeNext: false, 
          fixAlign: false
        }
      };
    });

    return { mod4Tables: tableInfos };
  }

  if (action === 'MODULE_4_FIX') {
    const { mod4File, mod4TableConfig } = payload;
    if (!mod4File || !mod4TableConfig) throw new Error("Thiếu dữ liệu.");

    const arrayBuffer = await mod4File.arrayBuffer();
    const pzip = new PizZip(arrayBuffer);
    const xmlStr = pzip.file("word/document.xml")?.asText() || "";
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlStr, "application/xml");
    const serializer = new XMLSerializer();

    const tblNodeMap = new Map<Element, Mod4TableInfo>();
    const initialDomTables = Array.from(doc.getElementsByTagName("w:tbl"));
    initialDomTables.forEach((tbl, i) => {
        const conf = mod4TableConfig.find(c => c.index === i);
        if (conf) tblNodeMap.set(tbl, conf);
    });

    // 1. Re-run Merge Logic
    const sortedConfigs = [...mod4TableConfig].sort((a, b) => b.index - a.index);
    
    for (const conf of sortedConfigs) {
       const tblNode = Array.from(tblNodeMap.keys()).find(k => tblNodeMap.get(k)?.id === conf.id);
       if (!tblNode || !tblNode.parentNode) continue; 

       if (conf.options.mergeNext) {
           let sibling = tblNode.nextSibling;
           const gapNodes: Node[] = [];
           let targetTbl: Element | null = null;
           while (sibling) {
               if (sibling.nodeName === 'w:tbl') { targetTbl = sibling as unknown as Element; break; }
               gapNodes.push(sibling);
               sibling = sibling.nextSibling;
           }

           if (targetTbl) {
               const rows = Array.from(targetTbl.getElementsByTagName("w:tr"));
               rows.forEach(r => tblNode.appendChild(r));
               gapNodes.forEach(n => n.parentNode?.removeChild(n));
               targetTbl.parentNode?.removeChild(targetTbl);
           }
       }
    }

    // 2. Apply Styles & Logic to remaining nodes
    for (const [node, conf] of tblNodeMap.entries()) {
        if (!node.parentNode) continue; // Skip removed nodes
        
        let tblPr = node.getElementsByTagName("w:tblPr")[0];
        if (!tblPr) {
            tblPr = doc.createElement("w:tblPr");
            node.insertBefore(tblPr, node.firstChild);
        }

        // A. Fix Borders (Ensure full borders)
        if (conf.options.fixBorders) {
            const existingBorders = tblPr.getElementsByTagName("w:tblBorders")[0];
            if (existingBorders) tblPr.removeChild(existingBorders);
            
            const newBorders = doc.createElement("w:tblBorders");
            ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].forEach(border => {
                const b = doc.createElement(`w:${border}`);
                b.setAttribute("w:val", "single");
                b.setAttribute("w:sz", "4"); // 1/2 pt
                b.setAttribute("w:space", "0");
                b.setAttribute("w:color", "auto");
                newBorders.appendChild(b);
            });
            tblPr.appendChild(newBorders);
        }

        // B. Autofit (85% width)
        if (conf.options.autofit) {
            let tblW = tblPr.getElementsByTagName("w:tblW")[0];
            if (!tblW) { tblW = doc.createElement("w:tblW"); tblPr.appendChild(tblW); }
            // 85% width. 5000 is 100%. 0.85 * 5000 = 4250.
            tblW.setAttribute("w:w", "4250");
            tblW.setAttribute("w:type", "pct");

            // Center the table
            let jc = tblPr.getElementsByTagName("w:jc")[0];
            if (!jc) { jc = doc.createElement("w:jc"); tblPr.appendChild(jc); }
            jc.setAttribute("w:val", "center");

            // Remove tblLayout to allow width to take precedence
            const layout = tblPr.getElementsByTagName("w:tblLayout")[0];
            if (layout) tblPr.removeChild(layout);
        }

        // C. Fix Cell Spacing & External Gaps
        if (conf.options.fixSpacing) {
             // 1. Internal Margins
             let cellMar = tblPr.getElementsByTagName("w:tblCellMar")[0];
             if (!cellMar) { cellMar = doc.createElement("w:tblCellMar"); tblPr.appendChild(cellMar); }
             const margins = { top: 50, bottom: 50, left: 100, right: 100 };
             Object.entries(margins).forEach(([side, val]) => {
                 let m = cellMar.getElementsByTagName(`w:${side}`)[0];
                 if (!m) { m = doc.createElement(`w:${side}`); cellMar.appendChild(m); }
                 m.setAttribute("w:w", String(val));
                 m.setAttribute("w:type", "dxa");
             });

             // 2. External Gaps (Keep 1 empty line)
             if (!conf.options.mergeNext) {
                let sibling = node.nextSibling;
                const gapNodes: Node[] = [];
                let nextTableFound = false;
                
                while(sibling) {
                    if (sibling.nodeName === 'w:tbl') {
                        nextTableFound = true;
                        break;
                    }
                    if (sibling.nodeName === 'w:p') {
                        const t = sibling.textContent || "";
                        if (t.trim() === "") {
                            gapNodes.push(sibling);
                        } else {
                            break; 
                        }
                    } else {
                        break;
                    }
                    sibling = sibling.nextSibling;
                }

                if (nextTableFound) {
                    if (gapNodes.length === 0) {
                        // Insert 1 empty paragraph if none exist
                        const p = doc.createElement("w:p");
                        node.parentNode?.insertBefore(p, node.nextSibling);
                    } else if (gapNodes.length > 1) {
                        // Remove all except the first one
                        for(let i=1; i < gapNodes.length; i++) {
                            gapNodes[i].parentNode?.removeChild(gapNodes[i]);
                        }
                    }
                }
             }
        }

        // D. Fix Align / Paragraph Spacing / Smart Bullets
        if (conf.options.fixAlign) {
            const cells = Array.from(node.getElementsByTagName("w:tc"));
            cells.forEach(cell => {
                // Clear cell width preference if autofit is on to let table handle width
                if (conf.options.autofit) {
                    const tcPr = cell.getElementsByTagName("w:tcPr")[0];
                    if (tcPr) {
                        const tcW = tcPr.getElementsByTagName("w:tcW")[0];
                        if (tcW) {
                             tcW.setAttribute("w:type", "auto");
                             tcW.setAttribute("w:w", "0");
                        }
                    }
                }

                const paras = Array.from(cell.getElementsByTagName("w:p"));
                paras.forEach(p => {
                    let pPr = p.getElementsByTagName("w:pPr")[0];
                    if (!pPr) { pPr = doc.createElement("w:pPr"); p.insertBefore(pPr, p.firstChild); }
                    
                    // Remove existing indentation
                    const ind = pPr.getElementsByTagName("w:ind")[0];
                    if (ind) pPr.removeChild(ind);

                    // Smart Bullet Detection (- + •)
                    const fullText = Array.from(p.getElementsByTagName("w:t")).map(t => t.textContent).join("").trim();
                    if (/^[-+•]/.test(fullText)) {
                        // Apply Hanging Indent (0.75cm ~ 425 dxa)
                        const newInd = doc.createElement("w:ind");
                        newInd.setAttribute("w:left", "425");
                        newInd.setAttribute("w:hanging", "425");
                        pPr.appendChild(newInd);
                    }

                    // Fix Paragraph Spacing
                    if (conf.options.fixSpacing) {
                        let spacing = pPr.getElementsByTagName("w:spacing")[0];
                        if (!spacing) { spacing = doc.createElement("w:spacing"); pPr.appendChild(spacing); }
                        spacing.setAttribute("w:before", "40"); 
                        spacing.setAttribute("w:after", "40"); 
                    }
                    
                    // Force Left Alignment
                    let jc = pPr.getElementsByTagName("w:jc")[0];
                    if (!jc) { jc = doc.createElement("w:jc"); pPr.appendChild(jc); }
                    jc.setAttribute("w:val", "left"); 
                });
            });
        }
    }

    // 4. Serialize and Save
    const newXml = serializer.serializeToString(doc);
    pzip.file("word/document.xml", newXml);
    
    const outBlob = pzip.generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
    
    const originalName = mod4File.name.substring(0, mod4File.name.lastIndexOf('.')) || "doc";
    return { blob: outBlob, filename: `${originalName}_fixed.docx` };
  }

  return { };
}

// ============================================================================
// MAIN CONTROLLER (CONSOLIDATED LOGIC)
// ============================================================================
async function main_controller(input: CoreInput, mode: number): Promise<CoreOutput> {
  // --- MODULE 1 LOGIC ---
  if (mode === 1) {
    const { action, payload } = input;

    // --- INTERNAL HELPER: Format Date Vietnamese ---
    const formatDateVI = (dateString: string) => {
      if (!dateString) return "";
      if (dateString.includes('/')) return dateString;
      if (/^\d{4}$/.test(dateString)) return dateString;
      try {
        const [y, m, d] = dateString.split('-');
        if (y && m && d) return `${d}/${m}/${y}`;
        return dateString;
      } catch (e) {
        return dateString;
      }
    };

    // --- INTERNAL HELPER: File to Base64 ---
    const fileToBase64 = (file: File): Promise<string> => {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result as string);
        reader.onerror = error => reject(error);
      });
    };

    // --- INTERNAL HELPER: Smart XML Filling ---
    const smartFillXML = (xml: string, info: StudentInfo, result: ESDMResult, levels: number[]): string => {
      try {
        const parser = new DOMParser();
        const doc = parser.parseFromString(xml, "application/xml");
        const serializer = new XMLSerializer();

        // 1. Fill Student Info & Summary
        const displayDob = formatDateVI(info.dob);
        const displayEvalDate = formatDateVI(info.evalDate);
        const percents = result.percents;
        const percentsOld = result.percentsOld;

        const infoMap = [
          { key: "Họ và tên học sinh", value: info.name },
          { key: "Họ và tên trẻ", value: info.name },
          { key: "Họ và tên", value: info.name },
          { key: "Họ tên", value: info.name },
          { key: "Tên trẻ", value: info.name },
          { key: "Ngày tháng năm sinh", value: displayDob },
          { key: "Ngày sinh", value: displayDob },
          { key: "Năm sinh", value: displayDob },
          { key: "Ngày lượng giá", value: displayEvalDate },
          { key: "Ngày đánh giá", value: displayEvalDate },
          { key: "Tuổi thực", value: info.age }, 
          { key: "Độ tuổi", value: info.age },
          { key: "Tuổi", value: info.age },
          { key: "Giới tính", value: info.gender },
          { key: "Mã học sinh", value: info.studentId },
          { key: "Mã HS", value: info.studentId },
          { key: "Mã số", value: info.studentId },
        ];

        const paragraphs = Array.from(doc.getElementsByTagName("w:p"));

        paragraphs.forEach(p => {
          const tNodes = Array.from(p.getElementsByTagName("w:t"));
          if (tNodes.length === 0) return;
          let fullText = "";
          tNodes.forEach(node => fullText += node.textContent || "");

          // Handle Summary (Red Percentages)
          if (fullText.toLowerCase().includes("nhận định chung về kết quả")) {
            let pPr = p.getElementsByTagName("w:pPr")[0];
            if (!pPr) {
              pPr = doc.createElement("w:pPr");
              if (p.firstChild) p.insertBefore(pPr, p.firstChild); else p.appendChild(pPr);
            }
            let jc = pPr.getElementsByTagName("w:jc")[0];
            if (!jc) {
              jc = doc.createElement("w:jc");
              pPr.appendChild(jc);
            }
            jc.setAttribute("w:val", "center");

            const regex = /(.*?Nhận định chung về kết quả.*?[:：])(.*?)(Tuổi phát triển.*|$)/i;
            const match = fullText.match(regex);
            const prefix = match ? match[1].trim() : "Nhận định chung về kết quả:";
            
            while (p.firstChild) {
              if (p.firstChild.nodeName !== "w:pPr") p.removeChild(p.firstChild);
              else {
                if (p.childNodes.length > 1) p.removeChild(p.childNodes[1]);
                else break; 
              }
            }

            const createRun = (text: string, isRed: boolean = false, isBold: boolean = false) => {
                const r = doc.createElement("w:r");
                const rPr = doc.createElement("w:rPr");
                const sz = doc.createElement("w:sz");
                sz.setAttribute("w:val", "26");
                rPr.appendChild(sz);
                const szCs = doc.createElement("w:szCs");
                szCs.setAttribute("w:val", "26");
                rPr.appendChild(szCs);
                if (isBold) rPr.appendChild(doc.createElement("w:b"));
                if (isRed) {
                    const color = doc.createElement("w:color");
                    color.setAttribute("w:val", "FF0000");
                    rPr.appendChild(color);
                }
                r.appendChild(rPr);
                const t = doc.createElement("w:t");
                t.setAttribute("xml:space", "preserve");
                t.textContent = text;
                r.appendChild(t);
                return r;
            };

            p.appendChild(createRun(prefix + " ", false, true)); 
            if (levels.length === 0) {
                p.appendChild(createRun(" Chưa chọn cấp độ nào để đánh giá."));
            } else {
                levels.forEach((l, index) => {
                    const valNew = (percents[`level${l}`] || 0).toFixed(1).replace('.', ',');
                    p.appendChild(createRun(`cấp độ ${l} con đạt `));
                    if (percentsOld && percentsOld[`level${l}`] !== undefined) {
                        const valOld = (percentsOld[`level${l}`] || 0).toFixed(1).replace('.', ',');
                        p.appendChild(createRun(`${valOld}% > ${valNew}%`, true, true));
                    } else {
                        p.appendChild(createRun(`${valNew}%`, true, true));
                    }
                    if (index < levels.length - 1) p.appendChild(createRun(". Và "));
                    else p.appendChild(createRun("."));
                });
            }
            return; 
          }

          // Handle General Info
          for (const { key, value } of infoMap) {
              if (!value) continue;
              const regex = new RegExp(`^\\s*${key}\\s*[:：]`, "i");
              if (regex.test(fullText)) {
                  const newText = `${key}: ${value}`;
                  tNodes[0].textContent = newText;
                  for(let i=1; i<tNodes.length; i++) tNodes[i].textContent = "";
                  break; 
              }
          }
        });

        // 2. Intelligent Table Filling
        const getCellText = (cell: Element): string => {
            const paragraphs = Array.from(cell.getElementsByTagName("w:p"));
            if (paragraphs.length === 0) return cell.textContent || "";
            return paragraphs.map(p => {
              return Array.from(p.getElementsByTagName("w:t")).map(t => t.textContent).join("");
            }).join(" ");
        };
        const normalize = (str: string) => str.replace(/\s+/g, ' ').trim().toLowerCase();
        const tables = Array.from(doc.getElementsByTagName("w:tbl"));
        
        for (const tbl of tables) {
            const rows = Array.from(tbl.getElementsByTagName("w:tr"));
            if (rows.length === 0) continue;
            let isESDMTable = false;
            let columnMap: Record<number, number> = {}; 
            
            for(let i=0; i < Math.min(5, rows.length); i++) {
                const cells = Array.from(rows[i].getElementsByTagName("w:tc"));
                if(cells.length === 0) continue;
                const firstCellText = normalize(getCellText(cells[0]));
                if (firstCellText.includes("kỹ năng")) {
                    isESDMTable = true;
                    for(let c=1; c < cells.length; c++) {
                        const cellText = normalize(getCellText(cells[c]));
                        const match = cellText.match(/(?:cấp độ|level)\s*(\d+)/);
                        if (match) {
                            const levelNum = parseInt(match[1]);
                            if (levelNum < 10) columnMap[levelNum] = c;
                            else {
                                const digitMatch = cellText.match(/(?:cấp độ|level)\s*(\d)/);
                                if (digitMatch) columnMap[parseInt(digitMatch[1])] = c;
                            }
                        }
                    }
                    break; 
                }
            }

            if (isESDMTable) {
                // Set Table Width to 90% (4500) and Center Align
                let tblPr = tbl.getElementsByTagName("w:tblPr")[0];
                if (!tblPr) {
                    tblPr = doc.createElement("w:tblPr");
                    tbl.insertBefore(tblPr, tbl.firstChild);
                }
                
                let tblW = tblPr.getElementsByTagName("w:tblW")[0];
                if (!tblW) {
                    tblW = doc.createElement("w:tblW");
                    tblPr.appendChild(tblW);
                }
                tblW.setAttribute("w:w", "4500");
                tblW.setAttribute("w:type", "pct");

                let jc = tblPr.getElementsByTagName("w:jc")[0];
                if (!jc) {
                    jc = doc.createElement("w:jc");
                    tblPr.appendChild(jc);
                }
                jc.setAttribute("w:val", "center");

                rows.forEach(row => {
                    const cells = Array.from(row.getElementsByTagName("w:tc"));
                    if (cells.length === 0) return;
                    const rowTitle = normalize(getCellText(cells[0]));
                    if (!rowTitle || rowTitle.includes("kỹ năng")) return;

                    const matchedSkill = result.table.find(item => {
                        const itemSkill = normalize(item.skill);
                        return itemSkill === rowTitle || rowTitle.startsWith(itemSkill);
                    });

                    if (matchedSkill) {
                        for (const [level, colIndex] of Object.entries(columnMap)) {
                            const lvl = parseInt(level);
                            if (levels.includes(lvl) && colIndex < cells.length) {
                                const cell = cells[colIndex];
                                let val = (matchedSkill as any)[`level${lvl}`];
                                if (val === undefined || val === null) val = "-";

                                const paragraphs = Array.from(cell.getElementsByTagName("w:p"));
                                if (paragraphs.length > 0) {
                                    const p = paragraphs[0];
                                    let r = p.getElementsByTagName("w:r")[0];
                                    if (!r) { r = doc.createElement("w:r"); p.appendChild(r); }
                                    let t = r.getElementsByTagName("w:t")[0];
                                    if (!t) { t = doc.createElement("w:t"); r.appendChild(t); }
                                    t.textContent = val;
                                    for(let k=1; k < paragraphs.length; k++) paragraphs[k].textContent = ""; 
                                    const runs = Array.from(p.getElementsByTagName("w:r"));
                                    for(let k=1; k < runs.length; k++) runs[k].textContent = "";
                                } else {
                                    const p = doc.createElement("w:p"); cell.appendChild(p);
                                    const r = doc.createElement("w:r"); p.appendChild(r);
                                    const t = doc.createElement("w:t"); t.textContent = val; r.appendChild(t);
                                }
                            }
                        }
                    }
                });
                break; 
            }
        }
        return serializer.serializeToString(doc);
      } catch (e) {
        console.warn("Smart Fill Warning:", e);
        return xml;
      }
    };

    switch (action) {
      case 'CALCULATE_AGE': {
        const { dob, evalDate, ageFormat } = payload;
        if (!dob || !evalDate) return { age: "" };

        const current = new Date(evalDate);
        let birthDate: Date | null = null;
        const dobTrim = dob.trim();
        
        const fullDateRegex = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/; 
        const monthYearRegex = /^(\d{1,2})[\/\-](\d{4})$/;
        const yearRegex = /^(\d{4})$/;
        const isoRegex = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;

        if (fullDateRegex.test(dobTrim)) {
          const [, d, m, y] = dobTrim.match(fullDateRegex)!;
          birthDate = new Date(parseInt(y), parseInt(m) - 1, parseInt(d));
        } else if (monthYearRegex.test(dobTrim)) {
          const [, m, y] = dobTrim.match(monthYearRegex)!;
          birthDate = new Date(parseInt(y), parseInt(m) - 1, 1);
        } else if (yearRegex.test(dobTrim)) {
          const [, y] = dobTrim.match(yearRegex)!;
          birthDate = new Date(parseInt(y), 0, 1);
        } else if (isoRegex.test(dobTrim)) {
          birthDate = new Date(dobTrim);
        }

        if (birthDate && !isNaN(birthDate.getTime()) && !isNaN(current.getTime())) {
          let years = current.getFullYear() - birthDate.getFullYear();
          let months = current.getMonth() - birthDate.getMonth();
          let days = current.getDate() - birthDate.getDate();

          if (days < 0) {
            months--;
            const prevMonth = new Date(current.getFullYear(), current.getMonth(), 0);
            days += prevMonth.getDate();
          }
          if (months < 0) {
            years--;
            months += 12;
          }

          if (years < 0) {
             return { age: "Ngày lượng giá nhỏ hơn ngày sinh" };
          } else {
             let ageStr = "";
             if (ageFormat === 'month') {
               const totalMonths = (years * 12) + months;
               ageStr = `${totalMonths} tháng`;
             } else {
               const parts = [];
               if (years > 0) parts.push(`${years} tuổi`);
               if (months > 0) parts.push(`${months} tháng`);
               if (days > 0 && years === 0) parts.push(`${days} ngày`);
               if (parts.length === 0) parts.push("0 tháng");
               ageStr = parts.join(" ");
             }
             return { age: ageStr };
          }
        }
        return { age: "" };
      }

      case 'ANALYZE': {
        const { files, selectedLevels, selectedColumns, apiKey, onProgress } = payload;
        if (!files || !selectedLevels || !selectedColumns) throw new Error("Missing params");

        const parts = [];
        const totalFiles = files.length;
        
        for (let i = 0; i < totalFiles; i++) {
          const file = files[i];
          const fileName = file.name.toLowerCase();
          
          if (file.type.startsWith('image/') || fileName.endsWith('.pdf')) {
            const base64 = await fileToBase64(file);
            parts.push({ inlineData: { mimeType: file.type || 'application/pdf', data: base64.split(',')[1] } });
          } else if (fileName.endsWith('.docx')) {
            const arrayBuffer = await file.arrayBuffer();
            const mammothResult = await mammoth.extractRawText({ arrayBuffer });
            parts.push({ text: mammothResult.value });
          }
          if (onProgress) onProgress(Math.round(((i + 1) / totalFiles) * 30));
        }

        const ai = new GoogleGenAI({ apiKey: apiKey || process.env.API_KEY });
        const levelsPrompt = selectedLevels.map(l => `CẤP ĐỘ ${l}`).join(', ');
        const sortedCols = [...selectedColumns].sort((a, b) => a - b);
        const isComparison = sortedCols.length > 1;
        
        let columnInstruction = "";
        if (!isComparison) {
          columnInstruction = `2. CHỈ đếm dấu "+" tại cột "Lần ${sortedCols[0]}". (Bỏ qua các cột khác). Trả về kết quả dạng "X/Y" (X là số đạt, Y là tổng).`;
        } else {
          columnInstruction = `2. Bạn cần so sánh 2 cột: Cột "Lần ${sortedCols[0]}" (Cũ) và Cột "Lần ${sortedCols[1]}" (Mới).
          - Đếm dấu "+" của cột "Lần ${sortedCols[0]}" (gọi là A).
          - Đếm dấu "+" của cột "Lần ${sortedCols[1]}" (gọi là B).
          - Trả về dữ liệu trong bảng dưới dạng chuỗi: "A/Total => B/Total". (Ví dụ: "2/4 => 4/4").
          - 'percents': tính % dựa trên cột MỚI NHẤT (Lần ${sortedCols[1]}).
          - 'percentsOld': tính % dựa trên cột CŨ HƠN (Lần ${sortedCols[0]}).`;
        }

        const prompt = `
Bạn là chuyên gia đánh giá ESDM chuyên sâu. Nhiệm vụ của bạn là đọc và trích xuất dữ liệu từ các trang của Phiếu Đánh Giá Chi Tiết ESDM.

QUY TẮC PHÂN TÍCH:
1. CHỈ xét các cấp độ: ${levelsPrompt}.
${columnInstruction}
3. Ký hiệu "+/-", "-", hoặc ô trống được tính là 0 mục đạt.
4. Mẫu số (tổng số mục) là tổng số dòng/mục con có trong danh sách kiểm tra của kỹ năng đó tại cấp độ đó.

PHẢI TRẢ VỀ DỮ LIỆU JSON CHÍNH XÁC VỚI CÁC TÊN KỸ NĂNG SAU:
- Giao tiếp tiếp nhận
- Giao tiếp diễn đạt
- Kỹ năng xã hội
- Bắt chước
- Nhận thức
- Chơi
- Vận động tinh
- Vận động thô
- Hành vi thích ứng
- Hành vi chú ý
- Tự lập
- Tổng điểm

Cấu trúc JSON:
{
  "table": [
    { "skill": "Tên kỹ năng", "level0": "...", "level1": "...", "level2": "...", "level3": "...", "level4": "..." },
    ...
  ],
  "percents": { "level0": float, "level1": float, "level2": float, "level3": float, "level4": float },
  ${isComparison ? '"percentsOld": { "level0": float, "level1": float, ... },' : ''}
  "summary": "Nhận xét tổng quát bằng tiếng Việt..."
}
`;

        const schemaProperties: any = {
          table: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                skill: { type: Type.STRING },
                level0: { type: Type.STRING },
                level1: { type: Type.STRING },
                level2: { type: Type.STRING },
                level3: { type: Type.STRING },
                level4: { type: Type.STRING }
              },
              required: ["skill", "level1", "level2", "level3", "level4"]
            }
          },
          percents: {
            type: Type.OBJECT,
            properties: {
              level0: { type: Type.NUMBER },
              level1: { type: Type.NUMBER },
              level2: { type: Type.NUMBER },
              level3: { type: Type.NUMBER },
              level4: { type: Type.NUMBER }
            }
          },
          summary: { type: Type.STRING }
        };

        if (isComparison) {
          schemaProperties.percentsOld = {
            type: Type.OBJECT,
            properties: {
              level0: { type: Type.NUMBER },
              level1: { type: Type.NUMBER },
              level2: { type: Type.NUMBER },
              level3: { type: Type.NUMBER },
              level4: { type: Type.NUMBER }
            }
          };
        }

        const response: GenerateContentResponse = await ai.models.generateContent({
          model: 'gemini-3-pro-preview',
          contents: { parts: [...parts, { text: prompt }] },
          config: {
            responseMimeType: "application/json",
            responseSchema: {
              type: Type.OBJECT,
              properties: schemaProperties,
              required: isComparison ? ["table", "percents", "percentsOld", "summary"] : ["table", "percents", "summary"]
            }
          }
        });

        let text = response.text || "{}";
        if (text.startsWith("```")) {
            text = text.replace(/^```json\s*/, "").replace(/^```\s*/, "").replace(/```$/, "");
        }
        return { esdmResult: JSON.parse(text) };
      }

      case 'GENERATE_DOCX': {
        const { templateFile, studentInfo, esdmResult, selectedLevels, fixCounter } = payload;
        if (!templateFile || !studentInfo || !esdmResult || !selectedLevels) throw new Error("Missing data");

        const arrayBuffer = await templateFile.arrayBuffer();
        const zip = new PizZip(arrayBuffer);

        const originalXml = zip.file("word/document.xml")?.asText();
        if (originalXml) {
          const smartXml = smartFillXML(originalXml, studentInfo, esdmResult, selectedLevels);
          zip.file("word/document.xml", smartXml);
        }

        const doc = new Docxtemplater(zip, {
          paragraphLoop: true,
          linebreaks: true,
          nullGetter: () => "" 
        });

        const displayDob = formatDateVI(studentInfo.dob);
        const displayEvalDate = formatDateVI(studentInfo.evalDate);

        const dataMapping: any = {
          name: studentInfo.name || "",
          dob: displayDob || "",
          eval_date: displayEvalDate || "",
          age: studentInfo.age || "",
          gender: studentInfo.gender || "",
          student_id: studentInfo.studentId || "",
          summary: "", 
          nhan_xet: "",
          p0: (esdmResult.percents.level0 || 0).toFixed(1) + '%',
          p1: (esdmResult.percents.level1 || 0).toFixed(1) + '%',
          p2: (esdmResult.percents.level2 || 0).toFixed(1) + '%',
          p3: (esdmResult.percents.level3 || 0).toFixed(1) + '%',
          p4: (esdmResult.percents.level4 || 0).toFixed(1) + '%',
        };

        doc.render(dataMapping);
        
        const out = doc.getZip().generate({
          type: 'blob',
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

        const originalName = templateFile.name.substring(0, templateFile.name.lastIndexOf('.')) || "File_Mau";
        const filename = `${originalName}_Fix${fixCounter || 1}.docx`;
        
        return { blob: out, filename };
      }
        
      default:
        throw new Error("Unknown action");
    }
  }

  // --- MODULE 2 LOGIC ---
  if (mode === 2) {
    if (input.action === 'PARSE_EXCEL') {
      const { file } = input.payload;
      if (!file) throw new Error("No file provided");
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);
      const levels: EsdmLevel[] = [];

      for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }) as any[][];
        if (!rows || rows.length === 0) continue;

        let smartGoalColIndex = -1;
        for (let r = 0; r < Math.min(rows.length, 30); r++) {
          const row = rows[r];
          if (!row) continue;
          const idx = row.findIndex(cell => {
            const val = String(cell || '').toLowerCase().trim();
            return val.includes('mục tiêu smart') || val === 'smart' || val === 'smart goal';
          });
          if (idx !== -1) { smartGoalColIndex = idx; break; }
        }
        
        const domains: EsdmDomain[] = [];
        let currentDomain: EsdmDomain | null = null;

        for (let i = 0; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;
          const val0 = String(row[0] || '').trim();
          if (!val0) continue;

          if (isGoalRowIndicator(val0)) {
            if (currentDomain) {
              let goalId = val0.replace(/\.$/, '').toUpperCase(); 
              if (/^\d+$/.test(goalId)) goalId = `M${goalId}`;
              let goalText = "";
              if (smartGoalColIndex !== -1 && row[smartGoalColIndex]) {
                 goalText = String(row[smartGoalColIndex]).trim();
              } else {
                 for(let k=1; k<row.length; k++) {
                    const cellVal = String(row[k] || '').trim();
                    if (cellVal.length > 5) { goalText = cellVal; break; }
                 }
              }
              if (goalText && !currentDomain.goals.find(g => g.id === goalId)) {
                currentDomain.goals.push({ id: goalId, text: goalText });
              }
            }
          } else if (isDomainLabel(val0)) {
            const domainName = val0;
            let existing = domains.find(d => d.name.toLowerCase() === domainName.toLowerCase());
            if (!existing) { existing = { name: domainName, goals: [] }; domains.push(existing); }
            currentDomain = existing;
          }
        }
        const validDomains = domains.filter(d => d.goals.length > 0);
        if (validDomains.length > 0) levels.push({ name: sheetName, domains: validDomains });
      }
      const sortedLevels = levels.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }));
      return { levelsData: sortedLevels };
    }

    if (input.action === 'GENERATE_IEP') {
      const { templateFile, originalFileName, selections, esdmData, smartSplitting } = input.payload;
      if (!templateFile || !selections || !esdmData) throw new Error("Missing params for IEP");

      // Prepare Data
      const processedData: ProcessedGoal[] = [];
      for (const sel of selections) {
        const levelData = esdmData.find(l => l.name === sel.level);
        const domainData = levelData?.domains.find(d => d.name === sel.domain);
        for (const sGoal of sel.goals) {
          const goal = domainData?.goals.find(g => g.id === sGoal.id);
          if (goal) processedData.push({ levelName: sel.level, domainName: sel.domain, goalId: goal.id, longTermGoal: goal.text, suffix: sGoal.suffix });
        }
      }

      const arrayBuffer = await templateFile.arrayBuffer();
      const zip = new PizZip(arrayBuffer);
      let documentXml = zip.file("word/document.xml")?.asText() || "";

      // Find Table
      let tableStartIndex = -1, tableEndIndex = -1, headerRowEndIndex = -1, found = false, hasSttColumn = false;
      let currentIndex = 0;
      while (true) {
        const tblStart = documentXml.indexOf('<w:tbl>', currentIndex);
        if (tblStart === -1) break;
        let depth = 1, scanIndex = tblStart + 7, tblEnd = -1;
        while (depth > 0) {
          const nextOpen = documentXml.indexOf('<w:tbl>', scanIndex);
          const nextClose = documentXml.indexOf('</w:tbl>', scanIndex);
          if (nextClose === -1) break; 
          if (nextOpen !== -1 && nextOpen < nextClose) { depth++; scanIndex = nextOpen + 7; } 
          else { depth--; scanIndex = nextClose + 8; if (depth === 0) tblEnd = nextClose + 8; }
        }
        if (tblEnd === -1) break;

        const firstRowEnd = documentXml.indexOf('</w:tr>', tblStart);
        if (firstRowEnd !== -1 && firstRowEnd < tblEnd) {
            const rawText = documentXml.substring(tblStart, firstRowEnd + 7).replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
            const normText = removeVietnameseTones(rawText);
            if (normText.includes("linh vuc") && normText.includes("muc tieu dai han") && normText.includes("muc tieu ngan han")) {
                if (normText.includes("stt") || normText.includes("no.")) hasSttColumn = true;
                tableStartIndex = tblStart; tableEndIndex = tblEnd; headerRowEndIndex = firstRowEnd + 7; found = true; break;
            }
        }
        currentIndex = tblEnd;
      }
      if (!found) throw new Error("Không tìm thấy bảng mục tiêu hợp lệ trong file Word.");

      // XML Builders
      const createCell = (content: string, opts: any = {}) => {
        let tcPr = opts.fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${opts.fill}"/>` : '';
        if (opts.gridSpan > 1) tcPr += `<w:gridSpan w:val="${opts.gridSpan}"/>`;
        if (opts.vMerge) tcPr += `<w:vMerge w:val="${opts.vMerge === 'restart' ? 'restart' : ''}"/>`;
        tcPr += `<w:vAlign w:val="${opts.bold ? 'center' : 'top'}"/>`;
        const pPr = `<w:pPr><w:jc w:val="${opts.align || 'left'}"/>${!opts.bold ? '<w:spacing w:after="100"/>' : ''}</w:pPr>`;
        let inner = opts.isXmlContent ? content : `<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/>${opts.bold ? '<w:b/><w:bCs/>' : ''}</w:rPr><w:t>${content}</w:t></w:r>`;
        return `<w:tc><w:tcPr>${tcPr}</w:tcPr><w:p>${pPr}${inner}</w:p></w:tc>`;
      };
      const createRow = (cells: string) => `<w:tr><w:trPr><w:trHeight w:val="400"/></w:trPr>${cells}</w:tr>`;

      let rowsXml = '';
      const domainsInOrder = Array.from(new Set(processedData.map(d => d.domainName)));
      
      domainsInOrder.forEach((domainName, dIdx) => {
        const domainGoals = processedData.filter(g => g.domainName === domainName);
        if (domainGoals.length === 0) return;
        const domainNum = dIdx + 1;
        
        // Logic Group Level: (CĐ1-M1, CĐ2-M5)
        const uniqueLevels = Array.from(new Set(domainGoals.map(g => g.levelName)));
        const levelParts = uniqueLevels.map(lvl => {
            const goals = domainGoals.filter(g => g.levelName === lvl);
            const shortLvl = (lvl.match(/\d+/) ? "CĐ" + lvl.match(/\d+/)?.[0] : lvl);
            return `${shortLvl}-${goals.map(g => g.goalId).join('-')}`;
        });
        const domainTitle = `${domainNum}. ${domainName} (${levelParts.join(', ')})`;
        
        domainGoals.forEach((goal, gIdx) => {
          const goalNum = `${domainNum}.${gIdx + 1}`;
          
          let shortGoalsRaw = [goal.longTermGoal, goal.longTermGoal, goal.longTermGoal];
          if (smartSplitting) {
              const txt = goal.longTermGoal;
              // 1. Range: a-b [unit]
              const rangeMatch = txt.match(/(\d{1,2})\s*-\s*(\d{1,2})(\s+(?:lần|bậc|loại|câu|từ|chữ))?/i);
              if (rangeMatch) {
                  const full = rangeMatch[0];
                  const a = parseInt(rangeMatch[1]);
                  const b = parseInt(rangeMatch[2]);
                  const suffix = rangeMatch[3] || "";
                  if (b > a) {
                      const a1 = Math.max(1, a - 2), b1 = Math.max(a1 + 1, b - 2); 
                      const a2 = Math.max(1, a - 1), b2 = Math.max(a2 + 1, b - 1);
                      shortGoalsRaw = [
                          txt.replace(full, `${a1}-${b1}${suffix}`),
                          txt.replace(full, `${a2}-${b2}${suffix}`),
                          txt
                      ];
                  }
              } else {
                // 2. Ratio: x/y
                const ratioMatch = txt.match(/(\d{1,3})\s*\/\s*(\d{1,3})/);
                if (ratioMatch) {
                     const full = ratioMatch[0];
                     const x = parseInt(ratioMatch[1]);
                     const y = parseInt(ratioMatch[2]);
                     if (x <= y) {
                         shortGoalsRaw = [
                             txt.replace(full, `${Math.max(1, x-2)}/${y}`),
                             txt.replace(full, `${Math.max(1, x-1)}/${y}`),
                             txt
                         ];
                     }
                } else {
                    // 3. Percent: x%
                    const pctMatch = txt.match(/(\d+)\s*%/);
                    if (pctMatch) {
                        const full = pctMatch[0];
                        const x = parseInt(pctMatch[1]);
                        shortGoalsRaw = [
                            txt.replace(full, `${Math.round(x/3)}%`),
                            txt.replace(full, `${Math.round(x*2/3)}%`),
                            txt
                        ];
                    } else {
                        // 4. Simple Count: x unit
                        const unitMatch = txt.match(/(\d+)\s+(lần|bậc|loại)/i);
                        if (unitMatch) {
                            const full = unitMatch[0];
                            const x = parseInt(unitMatch[1]);
                            const unit = unitMatch[2];
                            shortGoalsRaw = [
                                txt.replace(full, `${Math.max(1, x-2)} ${unit}`),
                                txt.replace(full, `${Math.max(1, x-1)} ${unit}`),
                                txt
                            ];
                        }
                    }
                }
              }
          }

          const longTermText = escapeXml(goal.longTermGoal);
          const suffix = escapeXml(goal.suffix || ''); 
          // Logic Bold Suffix
          const longTermXml = `<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>${goalNum}. ${longTermText}</w:t></w:r>` +
            (suffix ? `<w:r><w:br/></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="Times New Roman" w:cs="Times New Roman"/><w:b/><w:bCs/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t xml:space="preserve"> ${suffix}</w:t></w:r>` : '');

          for (let i = 1; i <= 3; i++) {
            const isDomStart = (gIdx === 0 && i === 1);
            const isGoalStart = (i === 1);
            let rowContent = "";
            const shortGoalEscaped = escapeXml(shortGoalsRaw[i-1]);
            if (hasSttColumn) rowContent += createCell(isDomStart ? String(domainNum) : "", { bold: true, vMerge: isDomStart ? 'restart' : 'continue', align: 'center' });
            rowContent += createCell(isDomStart ? escapeXml(domainTitle) : "", { bold: true, vMerge: isDomStart ? 'restart' : 'continue', align: 'left' });
            rowContent += createCell(isGoalStart ? longTermXml : "", { vMerge: isGoalStart ? 'restart' : 'continue', align: 'left', isXmlContent: true });
            rowContent += createCell(`${goalNum}.${i} ${shortGoalEscaped}`, { align: 'left' });
            rowsXml += createRow(rowContent);
          }
        });
      });

      const newDoc = documentXml.substring(0, headerRowEndIndex) + rowsXml + "</w:tbl>" + documentXml.substring(tableEndIndex);
      zip.file("word/document.xml", newDoc);
      const blob = zip.generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
      return { blob, filename: (originalFileName ? originalFileName.replace(/\.docx$/i, "") : "IEP") + "_fix1.docx" };
    }
    throw new Error("Unknown Action");
  }

  // --- MODULE 3 LOGIC ---
  if (mode === 3) {
    if (input.action === 'GENERATE_REPORT') {
      const { module3Data } = input.payload;
      if (!module3Data) throw new Error("Missing data for Module 3");

      const { childInfo, fieldGroups } = module3Data;

      // 1. Construct Prompt for JSON
      const prompt = `
DỮ LIỆU ĐẦU VÀO:
${JSON.stringify({ childInfo, fieldGroups })}

VAI TRÒ: Chuyên gia giáo dục đặc biệt (10 năm kinh nghiệm).
NHIỆM VỤ: Viết nội dung đánh giá và đề xuất chi tiết cho từng mục tiêu để điền vào báo cáo.

YÊU CẦU ĐẦU RA (JSON FORMAT):
Trả về một object JSON với cấu trúc sau:
{
  "goalSuggestions": [
    {
      "id": "string", // ID của mục tiêu từ dữ liệu đầu vào
      "assessment": "string", // Ví dụ: "+ CON HOÀN THÀNH 70% MỤC TIÊU ĐỀ RA." (Viết hoa toàn bộ, có dấu + ở đầu)
      "details": "string" // Nội dung đề xuất chi tiết (Sử dụng thẻ <b> cho các đầu mục, xuống dòng bằng <br/>)
    }
  ],
  "generalSummary": "string" // Đoạn văn tổng kết chung. Ví dụ: "Trong tháng 1, con hoàn thành 90% các hoạt động đề ra..."
}

YÊU CẦU CHI TIẾT NỘI DUNG "details":
- Bắt buộc có các đầu mục sau (in đậm bằng <b>...</b>:):
  <b>+Dạy trong bối cảnh thật:</b> ...
  <b>+Sử dụng đồ dùng hấp dẫn:</b> ...
  <b>+Kết hợp hành động - cử chỉ:</b> ...
  <b>+Giảm trợ giúp dần:</b> ...
  <b>+Lặp lại ở nhiều môi trường:</b> ...
  <b>+Khen khi đúng:</b> ...
- Nội dung cụ thể, thiết thực, giọng văn khích lệ.
`;

      // 2. Call Gemini
      const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });
      const response = await ai.models.generateContent({
        model: 'gemini-3-flash-preview',
        contents: { parts: [{ text: prompt }] },
        config: {
           responseMimeType: "application/json"
        }
      });

      const generatedContent = JSON.parse(response.text || "{}");
      const suggestionsMap = new Map<string, any>(generatedContent.goalSuggestions?.map((s: any) => [s.id, s]) || []);
      const generalSummary = generatedContent.generalSummary || "Chưa có tổng kết.";

      // 3. Generate DOCX using 'docx' library (creates real .docx file)
      
      // Helper: Parse simple HTML tags from Gemini (<b>, <br>) to Docx Paragraphs
      const parseHtmlToParagraphs = (text: string): Paragraph[] => {
          const lines = text.split(/<br\s*\/?>|\n/gi);
          return lines.filter(l => l.trim()).map(line => {
              const children: TextRun[] = [];
              const parts = line.split(/(<b>.*?<\/b>)/g);
              
              parts.forEach(part => {
                  if (part.startsWith('<b>') && part.endsWith('</b>')) {
                        children.push(new TextRun({
                            text: part.replace(/<\/?b>/g, ''),
                            bold: true,
                            font: "Times New Roman",
                            size: 26 // 13pt
                        }));
                  } else if (part) {
                        children.push(new TextRun({
                            text: part,
                            font: "Times New Roman",
                            size: 26
                        }));
                  }
              });
              
              return new Paragraph({
                  children: children,
                  spacing: { after: 100 }
              });
          });
      };

      const tableRows: TableRow[] = [];

      // --- Header Row 1 ---
      tableRows.push(
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: "LĨNH VỰC", bold: true, font: "Times New Roman", size: 26 })], alignment: AlignmentType.CENTER })],
              rowSpan: 2,
              shading: { fill: "70AD47", type: "clear", color: "auto" },
              verticalAlign: VerticalAlign.CENTER,
            }),
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: "MỤC TIÊU", bold: true, font: "Times New Roman", size: 26 })], alignment: AlignmentType.CENTER })],
              rowSpan: 2,
              shading: { fill: "70AD47", type: "clear", color: "auto" },
              verticalAlign: VerticalAlign.CENTER,
            }),
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: "KẾT QUẢ", bold: true, font: "Times New Roman", size: 26 })], alignment: AlignmentType.CENTER })],
              columnSpan: 3,
              shading: { fill: "70AD47", type: "clear", color: "auto" },
              verticalAlign: VerticalAlign.CENTER,
            }),
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: "ĐỀ XUẤT GIA ĐÌNH", bold: true, font: "Times New Roman", size: 26 })], alignment: AlignmentType.CENTER })],
              rowSpan: 2,
              shading: { fill: "70AD47", type: "clear", color: "auto" },
              verticalAlign: VerticalAlign.CENTER,
            }),
          ],
          // Explicitly NOT setting tableHeader: true prevents repeating header on page break
        })
      );

      // --- Header Row 2 ---
      tableRows.push(
        new TableRow({
           children: [
             new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "+", bold: true, font: "Times New Roman", size: 26 })], alignment: AlignmentType.CENTER })], shading: { fill: "70AD47", type: "clear", color: "auto" }, verticalAlign: VerticalAlign.CENTER }),
             new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "+/-", bold: true, font: "Times New Roman", size: 26 })], alignment: AlignmentType.CENTER })], shading: { fill: "70AD47", type: "clear", color: "auto" }, verticalAlign: VerticalAlign.CENTER }),
             new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "-", bold: true, font: "Times New Roman", size: 26 })], alignment: AlignmentType.CENTER })], shading: { fill: "70AD47", type: "clear", color: "auto" }, verticalAlign: VerticalAlign.CENTER }),
           ]
        })
      );

      // --- Data Rows ---
      fieldGroups.forEach((group, gIdx) => {
        group.goals.forEach((goal, i) => {
           const suggestionData = suggestionsMap.get(goal.id);
           const assessment = suggestionData?.assessment || `+ ĐẠT ${goal.percentage}% MỤC TIÊU.`;
           const details = suggestionData?.details || "";
           
           // Determine Mark
           let c1="", c2="", c3="";
           // Original logic: >=70 -> +, >=50 -> +/-, else -
           if (goal.percentage >= 70) c1 = "+";
           else if (goal.percentage >= 50) c2 = "+/-";
           else c3 = "-";

           const cells: TableCell[] = [];
           
           // Col 1: Field Name (RowSpan)
           if (i === 0) {
             cells.push(new TableCell({
               children: [new Paragraph({ children: [new TextRun({ text: `${gIdx + 1}. ${group.fieldName}`, bold: true, font: "Times New Roman", size: 26 })], alignment: AlignmentType.CENTER })],
               rowSpan: group.goals.length,
               verticalAlign: VerticalAlign.TOP,
             }));
           }

           // Col 2: Goal Text
           cells.push(new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: goal.goal, font: "Times New Roman", size: 26 })] })],
              verticalAlign: VerticalAlign.TOP,
           }));

           // Col 3, 4, 5: Marks
           [c1, c2, c3].forEach(mark => {
              cells.push(new TableCell({
                 children: [new Paragraph({ children: [new TextRun({ text: mark, font: "Times New Roman", size: 26 })], alignment: AlignmentType.CENTER })],
                 verticalAlign: VerticalAlign.CENTER,
              }));
           });

           // Col 6: Suggestions
           const suggestionParas = [
               new Paragraph({ children: [new TextRun({ text: assessment, bold: true, font: "Times New Roman", size: 26 })], spacing: { after: 100 } }),
               ...parseHtmlToParagraphs(details)
           ];
           if (goal.note) {
               suggestionParas.push(new Paragraph({ children: [new TextRun({ text: `(Ghi chú: ${goal.note})`, italics: true, font: "Times New Roman", size: 26 })], spacing: { before: 100 } }));
           }

           cells.push(new TableCell({
               children: suggestionParas,
               verticalAlign: VerticalAlign.TOP,
           }));

           tableRows.push(new TableRow({ children: cells }));
        });
      });

      // --- Summary Row ---
      tableRows.push(new TableRow({
        children: [
            new TableCell({
                children: [new Paragraph({ children: [new TextRun({ text: "TỔNG KẾT CHUNG", bold: true, font: "Times New Roman", size: 26 })], alignment: AlignmentType.CENTER })],
                shading: { fill: "FFC000", type: "clear", color: "auto" },
                verticalAlign: VerticalAlign.CENTER
            }),
            new TableCell({
                children: [new Paragraph({ children: [new TextRun({ text: generalSummary, font: "Times New Roman", size: 26 })] })],
                columnSpan: 5,
                shading: { fill: "FFC000", type: "clear", color: "auto" },
                verticalAlign: VerticalAlign.CENTER
            })
        ]
      }));

      // --- Document Creation ---
      const doc = new Document({
        sections: [{
          properties: {
             page: {
                margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } // 2cm ~ 1134 dxa
             }
          },
          children: [
            // Header Info
            new Paragraph({
                children: [
                    new TextRun({ text: "Trung Tâm Tâm lý-Giáo dục Sắc Màu", font: "Times New Roman", size: 26 }), // 13pt
                    new TextRun({ text: "\nĐịa chỉ: Lk 07, Ngõ 536a Minh Khai, Vĩnh Tuy, HBT, HN.", break: 1, font: "Times New Roman", size: 26 }),
                    new TextRun({ text: "\nLiên hệ: 0399797109", break: 1, font: "Times New Roman", size: 26 }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 400 }
            }),
            // Child Info Line 1
            new Paragraph({
                children: [
                    new TextRun({ text: "Họ và tên trẻ: ", bold: true, font: "Times New Roman", size: 26 }),
                    new TextRun({ text: childInfo.name, font: "Times New Roman", size: 26 }),
                    new TextRun({ text: "\t\tNgày sinh: ", bold: true, font: "Times New Roman", size: 26 }),
                    new TextRun({ text: childInfo.dob, font: "Times New Roman", size: 26 })
                ],
                tabStops: [{ type: "left", position: 6000 }],
                spacing: { after: 200 }
            }),
            // Child Info Line 2
            new Paragraph({
                children: [
                    new TextRun({ text: "Tháng báo cáo: ", bold: true, font: "Times New Roman", size: 26 }),
                    new TextRun({ text: childInfo.reportMonth, font: "Times New Roman", size: 26 })
                ],
                spacing: { after: 400 }
            }),
            // Table
            new Table({
                rows: tableRows,
                width: { size: 100, type: WidthType.PERCENTAGE },
                borders: {
                    top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                    bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                    left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                    right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                    insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                    insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
                },
                columnWidths: [1500, 2500, 500, 500, 500, 4500] // Approx weights matching percentages
            })
          ]
        }]
      });

      const blob = await Packer.toBlob(doc);
      const filename = `Bao_Cao_${childInfo.name.replace(/\s+/g, '_')}_${childInfo.reportMonth.replace(/\//g, '-')}.docx`;

      return { blob, filename };
    }
    throw new Error("Module 3: Unknown Action");
  }

  // --- MODULE 4 LOGIC ---
  if (mode === 4) {
    if (input.action === 'MODULE_4_ANALYZE' || input.action === 'MODULE_4_FIX') {
      return await autism_module_4(input);
    }
    throw new Error("Module 4: Unknown Action");
  }

  throw new Error(`Mode ${mode} không hợp lệ.`);
}

// --- MAIN COMPONENT (UI ONLY) ---
const App: React.FC = () => {
  // License State
  const [licenseState, setLicenseState] = useState<'checking' | 'verified' | 'unverified' | 'locked'>('checking');
  const [licenseMsg, setLicenseMsg] = useState('');
  const [inputKey, setInputKey] = useState('');

  // --- LICENSE VERIFICATION LOGIC ---
  useEffect(() => {
    const initSecurity = async () => {
        // Ensure Device ID
        let did = localStorage.getItem('deviceId');
        if (!did) {
            did = crypto.randomUUID();
            localStorage.setItem('deviceId', did);
        }
        
        // Check saved token
        const savedToken = localStorage.getItem('licenseToken');
        if (savedToken) {
            await verifyToken(savedToken);
        } else {
            setLicenseState('unverified');
        }
    };
    initSecurity();
  }, []);

  const verifyToken = async (token: string) => {
    setLicenseState('checking');
    setLicenseMsg('Đang xác thực bản quyền...');
    try {
        const did = localStorage.getItem('deviceId') || '';
        const fp = await generateAppFingerprint();
        const info = navigator.userAgent;
        
        // Call GAS API
        const url = `${API_LICENSE_URL}?token=${encodeURIComponent(token)}&deviceId=${encodeURIComponent(did)}&deviceInfo=${encodeURIComponent(info)}&fingerprint=${encodeURIComponent(fp)}`;
        
        const res = await fetch(url);
        const data = await res.json();
        
        if (data.ok) {
            localStorage.setItem('licenseToken', token);
            setLicenseState('verified');
        } else {
            setLicenseState('unverified');
            setLicenseMsg(data.message || 'Bản quyền không hợp lệ hoặc đã hết hạn.');
            if (data.message === 'Ứng dụng không hợp lệ') {
                 setLicenseState('locked');
                 setLicenseMsg('CẢNH BÁO BẢO MẬT: Phát hiện thay đổi mã nguồn hoặc thiết bị không hợp lệ. Vui lòng liên hệ quản trị viên.');
            }
        }
    } catch (e) {
        setLicenseState('unverified');
        setLicenseMsg('Lỗi kết nối máy chủ xác thực. Vui lòng kiểm tra mạng.');
    }
  };

  const handleLicenseSubmit = (e: React.FormEvent) => {
      e.preventDefault();
      if(inputKey.trim()) verifyToken(inputKey.trim());
  };

  // --- APP CONTENT IF VERIFIED ---
  const [appMode, setAppMode] = useState<number>(1); // 1: ESDM Standard, 2: Module 2, 3: Module 3, 4: Module 4
  
  // --- MODE 1 STATE ---
  const [files, setFiles] = useState<File[]>([]);
  const [previews, setPreviews] = useState<string[]>([]);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  
  const [status, setStatus] = useState<ProcessingStatus>(ProcessingStatus.IDLE);
  const [loadingProgress, setLoadingProgress] = useState<number>(0);
  const [loadingMessage, setLoadingMessage] = useState<string>("");
  
  const [result, setResult] = useState<ESDMResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [fixCounter, setFixCounter] = useState<number>(1);
  const [ageFormat, setAgeFormat] = useState<'detail' | 'month'>('detail');

  const [studentInfo, setStudentInfo] = useState<StudentInfo>({
    name: '',
    dob: '',
    evalDate: new Date().toISOString().split('T')[0],
    age: '',
    gender: 'Nam',
    studentId: ''
  });

  const [selectedLevels, setSelectedLevels] = useState<number[]>([]);
  const [selectedColumns, setSelectedColumns] = useState<number[]>([1]);

  // --- MODE 2 STATE ---
  const [esdmData, setEsdmData] = useState<EsdmLevel[]>([]);
  const [selections, setSelections] = useState<Selection[]>([]);
  const [mode2Template, setMode2Template] = useState<File | null>(null);
  const [expandedDomains, setExpandedDomains] = useState<Record<string, boolean>>({});
  const [loadingDefaultData, setLoadingDefaultData] = useState<boolean>(false);
  const [smartSplitting, setSmartSplitting] = useState<boolean>(false);

  // --- MODE 3 STATE ---
  const [mod3ChildInfo, setMod3ChildInfo] = useState<Mod3ChildInfo>({
    name: '', dob: '', reportMonth: '', caregiverTitle: 'bố mẹ'
  });
  const [mod3FieldGroups, setMod3FieldGroups] = useState<Mod3FieldGroup[]>([
    { id: '1', fieldName: 'Kỹ năng xã hội', goals: [{ id: '1-1', goal: '', percentage: 0, note: '' }] }
  ]);
  const [mod3Loading, setMod3Loading] = useState(false);

  // --- MODE 4 STATE ---
  const [mod4File, setMod4File] = useState<File | null>(null);
  const [mod4Tables, setMod4Tables] = useState<Mod4TableInfo[]>([]);
  const [mod4Loading, setMod4Loading] = useState(false);

  // Effect: Call controller to calculate age (Mode 1 Only mainly)
  useEffect(() => {
    if (licenseState === 'verified' && appMode === 1 && studentInfo.dob && studentInfo.evalDate) {
      main_controller({
        action: 'CALCULATE_AGE',
        payload: {
          dob: studentInfo.dob,
          evalDate: studentInfo.evalDate,
          ageFormat: ageFormat
        }
      }, 1).then(res => {
        if (res.age && res.age !== studentInfo.age) {
           setStudentInfo(prev => ({ ...prev, age: res.age! }));
        }
      }).catch(err => console.warn(err));
    }
  }, [studentInfo.dob, studentInfo.evalDate, ageFormat, appMode, licenseState]);

  // Effect: Auto-load default data for Mode 2
  useEffect(() => {
    if (licenseState === 'verified' && appMode === 2 && esdmData.length === 0) {
      handleLoadDefaultData();
    }
  }, [appMode, licenseState]);

  // --- MODE 1 HANDLERS ---
  const toggleLevel = (level: number) => {
    setSelectedLevels(prev => 
      prev.includes(level) ? prev.filter(l => l !== level) : [...prev, level].sort()
    );
  };

  const toggleColumn = (col: number) => {
    setSelectedColumns(prev => {
      if (prev.includes(col)) {
        const newVal = prev.filter(c => c !== col);
        return newVal.length === 0 ? [1] : newVal.sort((a, b) => a - b);
      } else {
        if (prev.length >= 2) {
          const [_, second] = prev;
          return [second, col].sort((a, b) => a - b);
        }
        return [...prev, col].sort((a, b) => a - b);
      }
    });
  };

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
    const { name, value } = e.target;
    setStudentInfo(prev => ({ ...prev, [name]: value }));
  };

  const handleDobChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    let value = e.target.value.replace(/\D/g, ''); 
    if (value.length > 8) value = value.slice(0, 8);
    let formatted = value;
    if (value.length > 4) formatted = `${value.slice(0, 2)}/${value.slice(2, 4)}/${value.slice(4)}`;
    else if (value.length > 2) formatted = `${value.slice(0, 2)}/${value.slice(2)}`;
    setStudentInfo(prev => ({ ...prev, dob: formatted }));
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      const selectedFiles = Array.from(e.target.files) as File[];
      selectedFiles.forEach(file => {
        const isPdf = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
        const isDocx = file.name.toLowerCase().endsWith('.docx');
        const isWord = file.type.includes('word') || file.name.toLowerCase().endsWith('.doc');

        if (isPdf || isDocx || isWord) {
          setFiles(prev => [...prev, file]);
          setPreviews(prev => [...prev, isPdf ? 'application/pdf' : (isDocx ? 'docx' : 'doc')]);
        } else if (file.type.startsWith('image/')) {
          setFiles(prev => [...prev, file]);
          const reader = new FileReader();
          reader.onloadend = () => setPreviews(prev => [...prev, reader.result as string]);
          reader.readAsDataURL(file);
        }
      });
    }
  };

  const handleTemplateUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) setTemplateFile(e.target.files[0]);
  };

  const removeFile = (index: number) => {
    setFiles(prev => prev.filter((_, i) => i !== index));
    setPreviews(prev => prev.filter((_, i) => i !== index));
  };

  const handleAnalyze = async () => {
    if (files.length === 0) return;
    setStatus(ProcessingStatus.LOADING);
    setLoadingProgress(0);
    setLoadingMessage("Đang đọc tài liệu...");
    setError(null);
    setResult(null);

    let progressInterval: any = setInterval(() => {
       setLoadingProgress(prev => {
         const next = prev + Math.random() * 2;
         return next > 90 ? 90 : Math.round(next);
       });
    }, 500);

    try {
      setLoadingMessage("AI đang phân tích & trích xuất...");
      const coreResponse = await main_controller({
        action: 'ANALYZE',
        payload: {
          files,
          selectedLevels,
          selectedColumns,
          apiKey: process.env.API_KEY,
          onProgress: (p) => setLoadingProgress(p)
        }
      }, 1);

      clearInterval(progressInterval);
      setLoadingProgress(100);
      setLoadingMessage("Hoàn tất!");
      
      setTimeout(() => {
        if (coreResponse.esdmResult) {
          setResult(coreResponse.esdmResult);
          setStatus(ProcessingStatus.SUCCESS);
        }
      }, 500);

    } catch (err: any) {
      if (progressInterval) clearInterval(progressInterval);
      setLoadingProgress(0);
      setError(err.message || "Lỗi hệ thống.");
      setStatus(ProcessingStatus.ERROR);
    }
  };

  const fillTemplate = async () => {
    if (!result || !templateFile) return;
    try {
      const coreResponse = await main_controller({
        action: 'GENERATE_DOCX',
        payload: { templateFile, studentInfo, esdmResult: result, selectedLevels, fixCounter }
      }, 1);

      if (coreResponse.blob && coreResponse.filename) {
        setFixCounter(prev => prev + 1);
        const url = URL.createObjectURL(coreResponse.blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = coreResponse.filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }
    } catch (err: any) {
      console.error("Template Error:", err);
      alert("Lỗi khi tạo file mới. Vui lòng kiểm tra file mẫu DOCX.");
    }
  };

  // --- MODE 2 HANDLERS ---
  const handleLoadDefaultData = async () => {
    setLoadingDefaultData(true);
    try {
      // Export URL for the Google Sheet
      const url = "https://docs.google.com/spreadsheets/d/1Os3f0967Po5wiJUEnCS51RhhikyMQcFD/export?format=xlsx";
      const response = await fetch(url);
      if (!response.ok) throw new Error("Không thể tải dữ liệu chuẩn.");
      
      const blob = await response.blob();
      const file = new File([blob], "ESDM_Data.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      
      const res = await main_controller({ action: 'PARSE_EXCEL', payload: { file } }, 2);
      if (res.levelsData) setEsdmData(res.levelsData);
    } catch (err) {
      console.error(err);
      alert("Lỗi tải dữ liệu ESDM chuẩn. Vui lòng thử lại sau.");
    } finally {
      setLoadingDefaultData(false);
    }
  };

  const toggleDomain = (domainKey: string) => {
    setExpandedDomains(prev => ({
      ...prev,
      [domainKey]: !prev[domainKey]
    }));
  };

  const handleMode2TemplateUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    if(e.target.files && e.target.files[0]) setMode2Template(e.target.files[0]);
  }

  const toggleGoalSelection = (levelName: string, domainName: string, goalId: string) => {
    setSelections(prev => {
      const existingSel = prev.find(s => s.level === levelName && s.domain === domainName);
      let newSelections = [...prev];

      if (existingSel) {
        const existingGoal = existingSel.goals.find(g => g.id === goalId);
        if (existingGoal) {
          // Remove goal
          const newGoals = existingSel.goals.filter(g => g.id !== goalId);
          if (newGoals.length === 0) {
            newSelections = newSelections.filter(s => !(s.level === levelName && s.domain === domainName));
          } else {
            newSelections = newSelections.map(s => s.level === levelName && s.domain === domainName ? { ...s, goals: newGoals } : s);
          }
        } else {
          // Add goal
          newSelections = newSelections.map(s => s.level === levelName && s.domain === domainName ? { ...s, goals: [...s.goals, { id: goalId, suffix: '(MTNT)' }] } : s);
        }
      } else {
        // Add new selection
        newSelections.push({ level: levelName, domain: domainName, goals: [{ id: goalId, suffix: '(MTNT)' }] });
      }
      return newSelections;
    });
  };

  const updateGoalSuffix = (levelName: string, domainName: string, goalId: string, suffix: GoalSuffix) => {
     setSelections(prev => prev.map(s => {
       if (s.level === levelName && s.domain === domainName) {
         return {
           ...s,
           goals: s.goals.map(g => g.id === goalId ? { ...g, suffix } : g)
         };
       }
       return s;
     }));
  };

  const generateIEP = async () => {
    if (!mode2Template || selections.length === 0) return;
    try {
      const res = await main_controller({
        action: 'GENERATE_IEP',
        payload: {
          templateFile: mode2Template,
          originalFileName: mode2Template.name,
          selections,
          esdmData,
          smartSplitting
        }
      }, 2);
      
      if (res.blob && res.filename) {
        const url = URL.createObjectURL(res.blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = res.filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
      }
    } catch(e) {
      console.error(e);
      alert("Lỗi tạo IEP.");
    }
  };

  // --- MODE 3 HANDLERS ---
  const handleMod3ChildChange = (e: React.ChangeEvent<HTMLInputElement | HTMLSelectElement>) => {
    const { name, value } = e.target;
    setMod3ChildInfo(prev => ({ ...prev, [name]: value }));
  };

  const addFieldGroup = () => {
    setMod3FieldGroups(prev => [...prev, { id: Date.now().toString(), fieldName: '', goals: [{ id: Date.now().toString() + 'g', goal: '', percentage: 0, note: '' }] }]);
  };

  const removeFieldGroup = (id: string) => {
    setMod3FieldGroups(prev => prev.filter(f => f.id !== id));
  };

  const updateFieldGroup = (id: string, name: string) => {
    setMod3FieldGroups(prev => prev.map(f => f.id === id ? { ...f, fieldName: name } : f));
  };

  const addGoalToField = (fieldId: string) => {
    setMod3FieldGroups(prev => prev.map(f => f.id === fieldId ? { ...f, goals: [...f.goals, { id: Date.now().toString(), goal: '', percentage: 0, note: '' }] } : f));
  };

  const removeGoal = (fieldId: string, goalId: string) => {
    setMod3FieldGroups(prev => prev.map(f => f.id === fieldId ? { ...f, goals: f.goals.filter(g => g.id !== goalId) } : f));
  };

  const updateGoal = (fieldId: string, goalId: string, field: keyof Mod3Goal, value: any) => {
    setMod3FieldGroups(prev => prev.map(f => {
      if (f.id === fieldId) {
        return {
          ...f,
          goals: f.goals.map(g => g.id === goalId ? { ...g, [field]: value } : g)
        };
      }
      return f;
    }));
  };

  const generateReport = async () => {
    setMod3Loading(true);
    try {
      const res = await main_controller({
        action: 'GENERATE_REPORT',
        payload: {
          module3Data: {
            childInfo: mod3ChildInfo,
            fieldGroups: mod3FieldGroups
          }
        }
      }, 3);

      if (res.blob && res.filename) {
        const url = URL.createObjectURL(res.blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = res.filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }
    } catch (e) {
      console.error(e);
      alert("Lỗi tạo báo cáo: " + (e as any).message);
    } finally {
      setMod3Loading(false);
    }
  };

  // --- MODE 4 HANDLERS ---
  const handleMod4FileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      const file = e.target.files[0];
      setMod4File(file);
      setMod4Loading(true);
      try {
        const res = await main_controller({
          action: 'MODULE_4_ANALYZE',
          payload: { mod4File: file }
        }, 4);
        if (res.mod4Tables) {
          setMod4Tables(res.mod4Tables);
        }
      } catch (err) {
        console.error(err);
        alert("Lỗi đọc file Word.");
      } finally {
        setMod4Loading(false);
      }
    }
  };

  const toggleMod4Option = (tableId: number, option: keyof Mod4TableInfo['options']) => {
    setMod4Tables(prev => prev.map(t => {
      if (t.id === tableId) {
        return {
          ...t,
          options: { ...t.options, [option]: !t.options[option] }
        };
      }
      return t;
    }));
  };

  const handleMod4Fix = async () => {
    if (!mod4File) return;
    setMod4Loading(true);
    try {
      const res = await main_controller({
        action: 'MODULE_4_FIX',
        payload: { mod4File, mod4TableConfig: mod4Tables }
      }, 4);

      if (res.blob && res.filename) {
        const url = URL.createObjectURL(res.blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = res.filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }
    } catch (err: any) {
      console.error(err);
      alert("Lỗi sửa file: " + err.message);
    } finally {
      setMod4Loading(false);
    }
  };

  // --- LICENSE GATE UI ---
  if (licenseState !== 'verified') {
    return (
      <div className="min-h-screen bg-slate-100 flex flex-col items-center justify-center p-4">
        <div className="bg-white p-8 rounded-3xl shadow-xl w-full max-w-md border border-slate-200">
          <div className="flex flex-col items-center mb-6">
            <div className="bg-indigo-600 p-4 rounded-2xl shadow-lg mb-4">
              <svg className="w-10 h-10 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z" /></svg>
            </div>
            <h1 className="text-2xl font-bold text-slate-800">Xác thực bản quyền</h1>
            <p className="text-sm text-slate-500 mt-2 text-center">Vui lòng nhập mã Token để kích hoạt ESDM Expert Pro.</p>
          </div>

          <form onSubmit={handleLicenseSubmit} className="space-y-4">
            <div>
              <label className="block text-xs font-bold text-slate-400 uppercase mb-1 ml-1">Mã Token</label>
              <input 
                type="text" 
                value={inputKey} 
                onChange={(e) => setInputKey(e.target.value)} 
                placeholder="Nhập mã token..." 
                className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all placeholder:text-slate-300 font-mono text-center tracking-widest text-lg uppercase"
                disabled={licenseState === 'checking' || licenseState === 'locked'}
              />
            </div>

            {licenseMsg && (
              <div className={`text-xs text-center font-medium p-3 rounded-lg ${licenseState === 'unverified' || licenseState === 'locked' ? 'bg-red-50 text-red-600' : 'bg-indigo-50 text-indigo-600'}`}>
                {licenseMsg}
              </div>
            )}

            <Button 
              type="submit" 
              disabled={licenseState === 'checking' || licenseState === 'locked' || !inputKey.trim()} 
              className="w-full h-12 text-base rounded-xl bg-indigo-600 hover:bg-indigo-700 shadow-indigo-200"
            >
              {licenseState === 'checking' ? 'Đang kiểm tra...' : 'Kích hoạt ngay'}
            </Button>
          </form>
          
          <div className="mt-6 pt-6 border-t border-slate-100 text-center">
             <p className="text-[10px] text-slate-400 uppercase tracking-wider font-bold">App ID: {APP_ID} | {BUILD_ID}</p>
          </div>
        </div>
      </div>
    );
  }

  // --- MAIN APP RENDER ---
  return (
    <div className="min-h-screen bg-slate-50 font-sans pb-20">
      <header className="bg-white border-b border-slate-200 sticky top-0 z-20 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="bg-indigo-600 p-2 rounded-lg shadow-indigo-200 shadow-lg">
              <svg className="w-6 h-6 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
            </div>
            <h1 className="text-xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-indigo-700 to-indigo-500">ESDM Expert v2.5.0</h1>
          </div>
          <div className="flex items-center gap-2">
            <select 
              value={appMode} 
              onChange={(e) => setAppMode(Number(e.target.value))}
              className="text-sm bg-slate-100 border border-slate-200 rounded-lg px-3 py-2 focus:outline-none focus:ring-2 focus:ring-indigo-500"
            >
              <option value={1}>Mode 1: Đánh giá ESDM</option>
              <option value={2}>Mode 2: Lập Kế Hoạch (IEP)</option>
              <option value={3}>Mode 3: Báo Cáo Can Thiệp</option>
              <option value={4}>Mode 4: Sửa Chữa Bảng (Fix Table)</option>
            </select>
            <Button onClick={() => window.location.reload()} variant="ghost" className="text-slate-500 hover:text-indigo-600">Làm mới</Button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 pt-8">
        
        {/* === MODE 1 UI === */}
        {appMode === 1 && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
            <div className="lg:col-span-4 space-y-6">
              <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100 ring-1 ring-slate-200/50">
                <h2 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
                  <span className="w-8 h-8 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center text-sm">1</span>
                  Thông tin trẻ
                </h2>
                <div className="grid gap-5">
                  <div>
                    <label className="text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-1 block">Họ và tên học sinh <span className="text-indigo-400 font-mono text-[10px]">{`{name}`}</span></label>
                    <input type="text" name="name" value={studentInfo.name} onChange={handleInputChange} placeholder="VD: Nguyễn Văn A" className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-2xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all placeholder:text-slate-300" />
                  </div>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label className="text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-1 block">Ngày sinh <span className="text-indigo-400 font-mono text-[10px]">{`{dob}`}</span></label>
                      <input type="text" name="dob" value={studentInfo.dob} onChange={handleDobChange} placeholder="dd/mm/yyyy" className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-2xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all placeholder:text-slate-300" />
                    </div>
                    <div>
                      <label className="text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-1 block">Ngày lượng giá <span className="text-indigo-400 font-mono text-[10px]">{`{eval_date}`}</span></label>
                      <input type="date" name="evalDate" value={studentInfo.evalDate} onChange={handleInputChange} className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-2xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all" />
                    </div>
                  </div>
                  <div className="grid grid-cols-2 gap-4">
                    <div>
                      <label className="text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-1 block">Giới tính <span className="text-indigo-400 font-mono text-[10px]">{`{gender}`}</span></label>
                      <select name="gender" value={studentInfo.gender} onChange={handleInputChange} className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-2xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all appearance-none cursor-pointer">
                        <option value="Nam">Nam</option>
                        <option value="Nữ">Nữ</option>
                      </select>
                    </div>
                    <div>
                      <div className="flex justify-between items-center mb-1">
                        <label className="text-[11px] font-bold text-slate-400 uppercase tracking-wider block">Tuổi thực <span className="text-indigo-400 font-mono text-[10px]">{`{age}`}</span></label>
                        <button onClick={() => setAgeFormat(prev => prev === 'detail' ? 'month' : 'detail')} className="text-[10px] text-indigo-600 hover:underline cursor-pointer font-medium">{ageFormat === 'detail' ? 'Đổi sang tháng' : 'Đổi sang chi tiết'}</button>
                      </div>
                      <input type="text" name="age" value={studentInfo.age} onChange={handleInputChange} placeholder="VD: 3 tuổi 2 tháng" className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-2xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all placeholder:text-slate-300" />
                    </div>
                  </div>
                  <div>
                    <label className="text-[11px] font-bold text-slate-400 uppercase tracking-wider mb-1 block">Mã học sinh <span className="text-indigo-400 font-mono text-[10px]">{`{student_id}`}</span></label>
                    <input type="text" name="studentId" value={studentInfo.studentId} onChange={handleInputChange} placeholder="VD: HS-102" className="w-full px-4 py-3 bg-slate-50 border border-slate-200 rounded-2xl focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500 outline-none transition-all placeholder:text-slate-300" />
                  </div>
                </div>
              </section>

              <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100 ring-1 ring-slate-200/50">
                <h2 className="text-lg font-bold text-slate-800 mb-6 flex items-center gap-2">
                  <span className="w-8 h-8 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center text-sm">2</span>
                  Cấp độ thống kê (+)
                </h2>
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                  {[0, 1, 2, 3, 4].map(l => (
                    <label key={l} className={`group relative flex items-center justify-center p-4 border-2 rounded-2xl cursor-pointer transition-all ${selectedLevels.includes(l) ? 'bg-indigo-600 border-indigo-600 text-white shadow-lg shadow-indigo-100' : 'bg-white border-slate-100 text-slate-400 hover:border-slate-200'}`}>
                      <input type="checkbox" checked={selectedLevels.includes(l)} onChange={() => toggleLevel(l)} className="hidden" />
                      <span className="font-bold">CĐ {l}</span>
                    </label>
                  ))}
                </div>
              </section>

              <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100 ring-1 ring-slate-200/50">
                <h2 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
                  <span className="w-8 h-8 rounded-full bg-indigo-100 text-indigo-600 flex items-center justify-center text-sm">3</span>
                  File kế hoạch (.docx)
                </h2>
                <div className="relative">
                  <input type="file" accept=".docx" onChange={handleTemplateUpload} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" />
                  <div className={`p-4 border-2 border-dashed rounded-2xl flex flex-col items-center justify-center text-center transition-all ${templateFile ? 'bg-emerald-50 border-emerald-200 text-emerald-600' : 'bg-slate-50 border-slate-200 text-slate-400 hover:bg-slate-100'}`}>
                    {templateFile ? (
                      <>
                        <svg className="w-8 h-8 mb-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" /></svg>
                        <span className="text-xs font-bold truncate w-full">{templateFile.name}</span>
                      </>
                    ) : (
                      <span className="text-[11px] text-slate-400 italic">Tải file mẫu Word (.docx) lên đây.</span>
                    )}
                  </div>
                </div>
              </section>
            </div>

            <div className="lg:col-span-8 space-y-6">
              <section className="bg-white p-8 rounded-3xl shadow-sm border border-slate-100 ring-1 ring-slate-200/50">
                <h2 className="text-xl font-bold text-slate-800 mb-2">Bảng kết quả ESDM</h2>
                <div className="mb-6 p-4 bg-slate-50 rounded-2xl border border-slate-100">
                  <div className="flex justify-between items-center mb-3">
                    <span className="text-xs font-bold text-slate-400 uppercase tracking-widest">Chọn cột phân tích (Max 2):</span>
                    <div className="text-xs font-medium text-slate-600 bg-white px-2 py-1 rounded border border-slate-200">
                      Đã chọn: {selectedColumns.map(c => `Lần ${c}`).join(" & ")}
                    </div>
                  </div>
                  <div className="grid grid-cols-9 gap-1 sm:gap-2">
                    {[1, 2, 3, 4, 5, 6, 7, 8, 9].map((col) => {
                      const isSelected = selectedColumns.includes(col);
                      const index = selectedColumns.indexOf(col);
                      return (
                        <button key={col} onClick={() => toggleColumn(col)} className={`aspect-square rounded-xl font-bold text-sm transition-all shadow-sm flex flex-col items-center justify-center relative ${isSelected ? 'bg-indigo-600 text-white shadow-indigo-200 scale-105 z-10' : 'bg-white text-slate-400 border border-slate-200'}`}>
                          {col}
                          {isSelected && <span className="absolute -top-2 -right-2 w-4 h-4 bg-orange-400 text-white rounded-full text-[8px] flex items-center justify-center border border-white">{index + 1}</span>}
                        </button>
                      );
                    })}
                  </div>
                </div>

                <div className="flex flex-wrap gap-4 mb-8">
                  {files.map((file, idx) => (
                    <div key={idx} className="relative group w-32 h-32 bg-slate-50 border border-slate-200 rounded-3xl overflow-hidden flex items-center justify-center shadow-sm">
                      {file.type.startsWith('image/') ? <img src={previews[idx]} className="w-full h-full object-cover" /> : <div className="flex flex-col items-center"><span className="text-[10px] font-black text-indigo-500 uppercase">{file.name.split('.').pop()}</span><span className="text-[8px] px-2 text-center text-slate-400 mt-1 truncate w-full">{file.name}</span></div>}
                      <button onClick={() => removeFile(idx)} className="absolute top-1 right-1 bg-white/90 text-red-500 rounded-full p-1.5 opacity-0 group-hover:opacity-100 transition-opacity border border-slate-100 shadow-sm"><svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M6 18L18 6M6 6l12 12" /></svg></button>
                    </div>
                  ))}
                  <label className="w-32 h-32 flex flex-col items-center justify-center border-2 border-dashed border-slate-200 rounded-3xl cursor-pointer hover:border-indigo-400 hover:bg-indigo-50/50 transition-all text-slate-300 hover:text-indigo-400 group">
                    <svg className="w-10 h-10 group-hover:scale-110 transition-transform" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M12 4v16m8-8H4" /></svg>
                    <span className="text-[10px] font-black uppercase mt-2">Thêm tài liệu</span>
                    <input type="file" multiple accept="image/*,application/pdf,.docx" onChange={handleFileChange} className="hidden" />
                  </label>
                </div>

                {status === ProcessingStatus.LOADING && <ProgressBar progress={loadingProgress} message={loadingMessage} />}
                <Button onClick={handleAnalyze} disabled={status === ProcessingStatus.LOADING || files.length === 0 || selectedLevels.length === 0 || selectedColumns.length === 0} className="w-full h-14 text-lg rounded-2xl shadow-xl shadow-indigo-100">
                  {status === ProcessingStatus.LOADING ? 'Đang xử lý...' : selectedLevels.length === 0 ? "Vui lòng chọn cấp độ" : `Bắt đầu thống kê`}
                </Button>
              </section>

              {status === ProcessingStatus.ERROR && <StatusAlert type="error" message={error || "Lỗi không xác định."} />}
              {result && (
                <div className="animate-in fade-in slide-in-from-bottom-8 duration-700 space-y-6">
                   <div className="flex justify-between items-center">
                     <h3 className="text-2xl font-black text-slate-800 tracking-tight">Kết quả</h3>
                     <Button onClick={fillTemplate} disabled={!templateFile} variant={templateFile ? 'primary' : 'secondary'} className={`h-12 px-8 rounded-2xl ${templateFile ? 'bg-emerald-600 hover:bg-emerald-700 shadow-emerald-100' : ''}`}>
                       {templateFile ? 'Tải Word' : 'Thiếu file mẫu'}
                     </Button>
                   </div>
                   <div className="bg-white rounded-3xl shadow-xl border border-slate-100 overflow-hidden ring-1 ring-slate-900/5">
                      <div className="p-8 border-b border-slate-100 bg-slate-50/50 flex justify-between">
                         <div><p className="text-[10px] font-black text-slate-400 uppercase tracking-widest mb-1">Kết quả</p><p className="text-xl font-bold text-slate-800">{studentInfo.name || '---'}</p></div>
                      </div>
                      <div className="overflow-x-auto">
                        <table className="w-full text-sm">
                          <thead><tr className="bg-slate-50 text-slate-500 border-b border-slate-100"><th className="p-4 text-left font-bold">Kỹ năng</th>{selectedLevels.map(l => <th key={l} className="p-4 text-center font-bold">CĐ {l}</th>)}</tr></thead>
                          <tbody>
                            {result.table.map((row, idx) => (
                              <tr key={idx} className={`border-b border-slate-50 hover:bg-indigo-50/20 transition-colors ${row.skill === "Tổng điểm" ? "font-bold bg-indigo-50/10 text-indigo-700" : "text-slate-600"}`}>
                                <td className="p-4">{row.skill}</td>
                                {selectedLevels.map(l => <td key={l} className="p-4 text-center font-mono font-medium">{(row as any)[`level${l}`]}</td>)}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      <div className="p-8 bg-slate-50/30">
                        <div className="p-6 bg-white rounded-2xl border border-slate-100 text-slate-600 text-sm leading-relaxed italic shadow-inner">{result.summary}</div>
                      </div>
                   </div>
                </div>
              )}
            </div>
          </div>
        )}

        {/* === MODE 2 UI === */}
        {appMode === 2 && (
          <div className="space-y-8">
            <div className="bg-white p-8 rounded-3xl shadow-sm border border-slate-100">
               <h2 className="text-xl font-bold text-slate-800 mb-6">1. Dữ liệu Mục tiêu (Excel)</h2>
               <div className="flex items-center gap-4 p-4 bg-slate-50 rounded-xl border border-slate-200">
                  {loadingDefaultData ? (
                    <div className="flex items-center text-indigo-600 font-medium">
                      <svg className="animate-spin h-5 w-5 mr-2" viewBox="0 0 24 24"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none"/><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>
                      Đang tải dữ liệu chuẩn từ hệ thống...
                    </div>
                  ) : esdmData.length > 0 ? (
                    <div className="flex items-center text-emerald-600 font-bold">
                      <svg className="w-6 h-6 mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M5 13l4 4L19 7" /></svg>
                      Dữ liệu chuẩn ESDM đã sẵn sàng.
                    </div>
                  ) : (
                    <div className="text-slate-500">Chưa có dữ liệu. Đang thử lại...</div>
                  )}
               </div>
            </div>

            {esdmData.length > 0 && (
               <div className="bg-white p-8 rounded-3xl shadow-sm border border-slate-100">
                 <div className="flex justify-between items-center mb-6">
                    <h2 className="text-xl font-bold text-slate-800">2. Chọn Mục tiêu Kế hoạch (IEP)</h2>
                    <label className="flex items-center gap-2 cursor-pointer bg-indigo-50 px-4 py-2 rounded-lg border border-indigo-100 hover:bg-indigo-100 transition-colors">
                      <input 
                        type="checkbox" 
                        checked={smartSplitting} 
                        onChange={(e) => setSmartSplitting(e.target.checked)} 
                        className="w-5 h-5 accent-indigo-600 rounded"
                      />
                      <span className="text-sm font-bold text-indigo-700">Mục tiêu nhỏ thông minh</span>
                    </label>
                 </div>
                 <div className="space-y-6">
                   {esdmData.map((level, idx) => (
                     <div key={idx} className="border border-slate-200 rounded-xl p-4">
                       <h3 className="font-bold text-lg text-indigo-600 mb-4">{level.name}</h3>
                       {level.domains.map((domain, dIdx) => {
                         const domainKey = `${level.name}-${domain.name}`;
                         const isExpanded = expandedDomains[domainKey];
                         const selectedCount = selections.find(s => s.level === level.name && s.domain === domain.name)?.goals.length || 0;

                         return (
                           <div key={dIdx} className="mb-3 last:mb-0">
                              <div className="flex items-center justify-between bg-slate-50 p-3 rounded-lg border border-slate-100 hover:border-indigo-200 transition-colors cursor-pointer" onClick={() => toggleDomain(domainKey)}>
                                <div className="flex items-center gap-3">
                                  <button className={`w-6 h-6 flex items-center justify-center rounded-full text-white font-bold text-xs transition-colors ${isExpanded ? 'bg-indigo-500' : 'bg-slate-300'}`}>
                                    {isExpanded ? '-' : '+'}
                                  </button>
                                  <h4 className="font-semibold text-slate-700">{domain.name}</h4>
                                </div>
                                {selectedCount > 0 && (
                                  <span className="bg-indigo-100 text-indigo-700 text-xs font-bold px-2 py-1 rounded-full">
                                    Đã chọn: {selectedCount}
                                  </span>
                                )}
                              </div>
                              
                              {isExpanded && (
                                <div className="mt-2 pl-4 grid grid-cols-1 md:grid-cols-2 gap-2 animate-in slide-in-from-top-2 fade-in duration-200">
                                   {domain.goals.map((goal, gIdx) => {
                                     const isSelected = selections.some(s => s.level === level.name && s.domain === domain.name && s.goals.some(g => g.id === goal.id));
                                     const currentSuffix = selections.find(s => s.level === level.name && s.domain === domain.name)?.goals.find(g => g.id === goal.id)?.suffix || '(MTNT)';
                                     
                                     return (
                                       <div key={gIdx} className={`p-2 rounded border transition-colors ${isSelected ? 'bg-indigo-50 border-indigo-200' : 'bg-white border-slate-100'}`}>
                                          <div className="flex items-start gap-2">
                                            <input 
                                              type="checkbox" 
                                              checked={isSelected} 
                                              onChange={() => toggleGoalSelection(level.name, domain.name, goal.id)}
                                              className="mt-1 accent-indigo-600"
                                            />
                                            <div className="flex-1">
                                               <p className="text-sm font-medium text-slate-800"><span className="text-indigo-500 font-bold mr-1">{goal.id}</span> {goal.text}</p>
                                               {isSelected && (
                                                  <select 
                                                    value={currentSuffix} 
                                                    onChange={(e) => updateGoalSuffix(level.name, domain.name, goal.id, e.target.value as GoalSuffix)}
                                                    className="mt-2 text-xs border border-slate-300 rounded px-2 py-1 bg-white focus:border-indigo-500 outline-none w-full"
                                                  >
                                                    <option value="(MTNT)">Ngắn hạn (MTNT)</option>
                                                    <option value="(MTC)">Chung (MTC)</option>
                                                    <option value="(MTP)">Phụ (MTP)</option>
                                                  </select>
                                               )}
                                            </div>
                                          </div>
                                       </div>
                                     )
                                   })}
                                </div>
                              )}
                           </div>
                         )
                       })}
                     </div>
                   ))}
                 </div>
               </div>
            )}

            <div className="bg-white p-8 rounded-3xl shadow-sm border border-slate-100">
               <h2 className="text-xl font-bold text-slate-800 mb-6">3. Tải Template & Xuất Word</h2>
               <div className="flex flex-col gap-4">
                  <div className="flex items-center gap-4">
                      <input type="file" accept=".docx" onChange={handleMode2TemplateUpload} className="block w-full text-sm text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-emerald-50 file:text-emerald-700 hover:file:bg-emerald-100"/>
                      {mode2Template && <span className="text-emerald-600 font-bold">Template: {mode2Template.name}</span>}
                  </div>
                  <Button 
                    onClick={generateIEP} 
                    disabled={!mode2Template || selections.length === 0}
                    className="h-12 text-lg w-full md:w-auto self-start"
                  >
                     Tạo File Kế Hoạch (IEP)
                  </Button>
               </div>
            </div>
          </div>
        )}

        {/* === MODE 3 UI === */}
        {appMode === 3 && (
           <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
              {/* CHILD INFO */}
              <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
                <h2 className="text-lg font-bold text-slate-800 mb-4">Thông tin trẻ & Báo cáo</h2>
                <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
                  <div>
                    <label className="text-xs font-bold text-slate-400 uppercase">Tên trẻ</label>
                    <input type="text" name="name" value={mod3ChildInfo.name} onChange={handleMod3ChildChange} className="w-full mt-1 px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg"/>
                  </div>
                  <div>
                    <label className="text-xs font-bold text-slate-400 uppercase">Ngày sinh</label>
                    <input type="text" name="dob" value={mod3ChildInfo.dob} onChange={handleMod3ChildChange} placeholder="dd/mm/yyyy" className="w-full mt-1 px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg"/>
                  </div>
                  <div>
                    <label className="text-xs font-bold text-slate-400 uppercase">Tháng báo cáo</label>
                    <input type="text" name="reportMonth" value={mod3ChildInfo.reportMonth} onChange={handleMod3ChildChange} placeholder="12/2023" className="w-full mt-1 px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg"/>
                  </div>
                  <div>
                    <label className="text-xs font-bold text-slate-400 uppercase">Gửi tới (Danh xưng)</label>
                    <select name="caregiverTitle" value={mod3ChildInfo.caregiverTitle} onChange={handleMod3ChildChange} className="w-full mt-1 px-3 py-2 bg-slate-50 border border-slate-200 rounded-lg">
                      <option value="bố">Bố</option>
                      <option value="mẹ">Mẹ</option>
                      <option value="bố mẹ">Bố Mẹ</option>
                    </select>
                  </div>
                </div>
              </div>

              {/* FIELDS & GOALS */}
              <div className="space-y-4">
                {mod3FieldGroups.map((group, idx) => (
                  <div key={group.id} className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100 relative group/card">
                    <button onClick={() => removeFieldGroup(group.id)} className="absolute top-4 right-4 text-slate-300 hover:text-red-500 opacity-0 group-hover/card:opacity-100 transition-opacity">
                      <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M6 18L18 6M6 6l12 12" strokeWidth="2"/></svg>
                    </button>
                    
                    <div className="mb-4">
                      <label className="text-xs font-bold text-indigo-500 uppercase mb-1 block">Lĩnh vực {idx + 1}</label>
                      <input 
                        type="text" 
                        value={group.fieldName} 
                        onChange={(e) => updateFieldGroup(group.id, e.target.value)} 
                        placeholder="VD: Kỹ năng xã hội" 
                        className="w-full text-lg font-bold text-slate-800 border-b border-slate-200 focus:border-indigo-500 outline-none py-1 placeholder:text-slate-300"
                      />
                    </div>

                    <div className="space-y-3 pl-4 border-l-2 border-slate-100">
                      {group.goals.map((goal) => (
                        <div key={goal.id} className="grid grid-cols-12 gap-4 items-center group/goal mb-4 bg-slate-50 p-2 rounded-lg">
                           <div className="col-span-12 md:col-span-5">
                             <label className="text-[10px] font-bold text-slate-400 uppercase block mb-1">Mục tiêu</label>
                             <input type="text" value={goal.goal} onChange={(e) => updateGoal(group.id, goal.id, 'goal', e.target.value)} placeholder="Nhập mục tiêu..." className="w-full text-sm px-2 py-2 bg-white rounded border border-slate-200 focus:border-indigo-200 outline-none"/>
                           </div>
                           <div className="col-span-12 md:col-span-3">
                             <div className="flex justify-between mb-1">
                               <label className="text-[10px] font-bold text-slate-400 uppercase">Mức độ đạt</label>
                               <span className="text-[10px] font-bold text-indigo-600">{goal.percentage}%</span>
                             </div>
                             <input 
                               type="range" 
                               min="0" 
                               max="100" 
                               value={goal.percentage} 
                               onChange={(e) => updateGoal(group.id, goal.id, 'percentage', parseInt(e.target.value))} 
                               className="w-full h-2 bg-slate-200 rounded-lg appearance-none cursor-pointer accent-indigo-600"
                             />
                           </div>
                           <div className="col-span-11 md:col-span-3">
                             <label className="text-[10px] font-bold text-slate-400 uppercase block mb-1">Ghi chú / Nhận xét</label>
                             <input type="text" value={goal.note} onChange={(e) => updateGoal(group.id, goal.id, 'note', e.target.value)} placeholder="Chi tiết..." className="w-full text-sm px-2 py-2 bg-white rounded border border-slate-200 focus:border-indigo-200 outline-none italic text-slate-600"/>
                           </div>
                           <div className="col-span-1 text-center flex items-end justify-center h-full pb-2">
                             <button onClick={() => removeGoal(group.id, goal.id)} className="text-slate-300 hover:text-red-400 p-1"><svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" strokeLinecap="round" strokeLinejoin="round" strokeWidth="2"/></svg></button>
                           </div>
                        </div>
                      ))}
                      <button onClick={() => addGoalToField(group.id)} className="text-xs font-bold text-indigo-600 hover:text-indigo-700 flex items-center gap-1 mt-2">
                        <svg className="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M12 4v16m8-8H4"/></svg> Thêm mục tiêu
                      </button>
                    </div>
                  </div>
                ))}
                
                <button onClick={addFieldGroup} className="w-full py-4 border-2 border-dashed border-slate-200 rounded-2xl text-slate-400 font-bold hover:border-indigo-300 hover:text-indigo-500 hover:bg-indigo-50 transition-all flex items-center justify-center gap-2">
                  <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M12 4v16m8-8H4" strokeWidth="2"/></svg>
                  Thêm Lĩnh vực mới
                </button>
              </div>

              {/* ACTION */}
              <div className="flex justify-end pt-4">
                <Button 
                  onClick={generateReport} 
                  disabled={mod3Loading}
                  className="h-14 px-8 text-lg shadow-xl shadow-indigo-100 bg-indigo-600 hover:bg-indigo-700 text-white rounded-2xl"
                >
                  {mod3Loading ? (
                    <>
                      <svg className="animate-spin h-5 w-5 mr-2" viewBox="0 0 24 24"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none"/><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>
                      Đang tạo báo cáo...
                    </>
                  ) : "Tạo Báo Cáo (Word)"}
                </Button>
              </div>
           </div>
        )}

        {/* === MODE 4 UI === */}
        {appMode === 4 && (
          <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-500">
            <div className="bg-white p-8 rounded-3xl shadow-sm border border-slate-100">
              <h2 className="text-xl font-bold text-slate-800 mb-2">Sửa Chữa Bảng Chuyên Sâu</h2>
              <p className="text-sm text-slate-500 mb-6">Tự động phát hiện và sửa các lỗi bảng biểu trong file Word mà không làm thay đổi nội dung văn bản.</p>
              
              <div className="flex items-center gap-4">
                <label className="flex-1 cursor-pointer">
                  <div className="flex flex-col items-center justify-center w-full h-32 border-2 border-dashed border-slate-300 rounded-2xl bg-slate-50 hover:bg-indigo-50 hover:border-indigo-300 transition-all group">
                    {mod4File ? (
                      <div className="flex flex-col items-center">
                        <svg className="w-8 h-8 text-emerald-500 mb-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z" strokeWidth="2"/></svg>
                        <span className="font-bold text-slate-700">{mod4File.name}</span>
                      </div>
                    ) : (
                      <>
                        <svg className="w-8 h-8 text-slate-400 group-hover:text-indigo-500 mb-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12" strokeWidth="2"/></svg>
                        <span className="text-sm text-slate-500">Tải file Word (.docx)</span>
                      </>
                    )}
                    <input type="file" accept=".docx" onChange={handleMod4FileChange} className="hidden" />
                  </div>
                </label>
                {mod4File && (
                  <Button 
                    onClick={handleMod4Fix} 
                    disabled={mod4Loading || mod4Tables.length === 0}
                    className="h-32 px-8 text-lg rounded-2xl shadow-xl shadow-indigo-100 bg-indigo-600 hover:bg-indigo-700 text-white"
                  >
                    {mod4Loading ? 'Đang xử lý...' : 'Sửa & Tải xuống'}
                  </Button>
                )}
              </div>
            </div>

            {mod4Tables.length > 0 && (
              <div className="grid gap-6">
                <div className="flex justify-between items-center">
                  <h3 className="font-bold text-lg text-slate-700">Danh sách bảng ({mod4Tables.length})</h3>
                  <div className="flex gap-2">
                    <span className="text-xs bg-red-100 text-red-600 px-2 py-1 rounded font-bold">Cảnh báo lỗi</span>
                    <span className="text-xs bg-indigo-100 text-indigo-600 px-2 py-1 rounded font-bold">Kiến nghị sửa</span>
                  </div>
                </div>
                
                {mod4Tables.map((tbl) => (
                  <div key={tbl.id} className={`bg-white rounded-2xl shadow-sm border overflow-hidden transition-all ${tbl.issues.length > 0 ? 'border-red-200 ring-2 ring-red-50' : 'border-slate-100'}`}>
                    <div className="p-4 bg-slate-50 border-b border-slate-100 flex justify-between items-center">
                      <div className="flex items-center gap-3">
                        <span className="w-8 h-8 rounded-full bg-white border border-slate-200 flex items-center justify-center font-bold text-sm text-slate-600">
                          {tbl.index + 1}
                        </span>
                        <div>
                          <h4 className="font-bold text-slate-700">Bảng số {tbl.index + 1}</h4>
                          {tbl.issues.length > 0 ? (
                            <div className="flex gap-2 mt-1">
                              {tbl.issues.map((issue, i) => (
                                <span key={i} className="text-[10px] font-bold uppercase tracking-wider text-red-500 bg-red-50 px-2 py-0.5 rounded border border-red-100">{issue}</span>
                              ))}
                            </div>
                          ) : <span className="text-xs text-emerald-600 font-medium">Bảng ổn định</span>}
                        </div>
                      </div>
                      
                      <div className="flex gap-4">
                        <label className="flex items-center gap-2 cursor-pointer">
                          <input type="checkbox" checked={tbl.options.fixBorders} onChange={() => toggleMod4Option(tbl.id, 'fixBorders')} className="w-4 h-4 rounded text-indigo-600 focus:ring-indigo-500 border-slate-300" />
                          <span className="text-sm font-medium text-slate-600">Đầy đủ viền</span>
                        </label>
                        <label className="flex items-center gap-2 cursor-pointer">
                          <input type="checkbox" checked={tbl.options.autofit} onChange={() => toggleMod4Option(tbl.id, 'autofit')} className="w-4 h-4 rounded text-indigo-600 focus:ring-indigo-500 border-slate-300" />
                          <span className="text-sm font-medium text-slate-600">Autofit (85%)</span>
                        </label>
                        <label className="flex items-center gap-2 cursor-pointer">
                          <input type="checkbox" checked={tbl.options.fixSpacing} onChange={() => toggleMod4Option(tbl.id, 'fixSpacing')} className="w-4 h-4 rounded text-indigo-600 focus:ring-indigo-500 border-slate-300" />
                          <span className="text-sm font-medium text-slate-600">Xoá khoảng trống</span>
                        </label>
                        <label className="flex items-center gap-2 cursor-pointer">
                          <input type="checkbox" checked={tbl.options.fixAlign} onChange={() => toggleMod4Option(tbl.id, 'fixAlign')} className="w-4 h-4 rounded text-indigo-600 focus:ring-indigo-500 border-slate-300" />
                          <span className="text-sm font-medium text-slate-600">Căn chỉnh chữ</span>
                        </label>
                      </div>
                    </div>

                    <div className="grid grid-cols-12">
                      <div className="col-span-8 p-6 bg-slate-50/50 border-r border-slate-100 overflow-x-auto">
                        <div className="min-w-full">
                          <div 
                            className="preview-table-container text-xs scale-90 origin-top-left opacity-75 pointer-events-none"
                            dangerouslySetInnerHTML={{ __html: tbl.previewHtml }} 
                          />
                        </div>
                      </div>
                      <div className="col-span-4 p-6 bg-white space-y-4">
                        <h5 className="font-bold text-xs text-slate-400 uppercase tracking-widest">Hành động nâng cao</h5>
                        
                        {tbl.canMergeNext ? (
                          <div className="p-3 bg-indigo-50 rounded-xl border border-indigo-100">
                            <label className="flex items-start gap-3 cursor-pointer">
                              <input type="checkbox" checked={tbl.options.mergeNext} onChange={() => toggleMod4Option(tbl.id, 'mergeNext')} className="mt-1 w-4 h-4 rounded text-indigo-600 focus:ring-indigo-500 border-indigo-300" />
                              <div>
                                <span className="block font-bold text-indigo-700 text-sm">Gộp với bảng kế tiếp</span>
                                <span className="block text-xs text-indigo-500 mt-1">Phát hiện bảng bị ngắt quãng. Chọn để nối liền bảng này với bảng số {tbl.index + 2}.</span>
                              </div>
                            </label>
                          </div>
                        ) : (
                          <div className="p-3 bg-slate-50 rounded-xl border border-slate-100 opacity-50 cursor-not-allowed">
                            <span className="block font-bold text-slate-400 text-sm">Không thể gộp bảng</span>
                            <span className="block text-xs text-slate-400 mt-1">Khoảng cách quá xa hoặc không liên tiếp.</span>
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
};

export default App;
