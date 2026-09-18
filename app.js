const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL_PACKAGE_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const WPS_NS = "http://www.wps.cn/officeDocument/2017/etCustomData";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const DRAWING_MAIN_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const GROUP_ORDER = ["上午", "中午", "下午"];
const anchors = { 上午: "A4", 中午: "A6", 下午: "A8" };
const OUTPUT_IMAGE_WIDTH = 1706;
const OUTPUT_IMAGE_HEIGHT = 1279;

const photosInput = document.querySelector("#photos");
const secondRowInput = document.querySelector("#second-row");
const fileSummary = document.querySelector("#file-summary");
const generateButton = document.querySelector("#generate");
const statusBox = document.querySelector("#status");
const downloadFallback = document.querySelector("#download-fallback");
let currentDownloadUrl = "";

photosInput.addEventListener("change", () => {
  const count = photosInput.files.length;
  fileSummary.textContent = count ? `已选择 ${count} 张照片` : "还没有选择照片";
});

function setStatus(message, type = "") {
  statusBox.textContent = message;
  statusBox.className = `status ${type}`.trim();
}

function groupName(hour) {
  if (hour >= 5 && hour <= 9) return "上午";
  if (hour >= 10 && hour <= 14) return "中午";
  if (hour >= 15) return "下午";
  return "未分类";
}

function filenameDate(file) {
  const pattern = /(20\d{2})[-_/.]?(\d{2})[-_/.]?(\d{2})[ _-]?(\d{2})[-_.]?(\d{2})(?:[-_.]?(\d{2}))?/;
  const match = file.name.match(pattern);
  if (!match) return null;
  const [, year, month, day, hour, minute, second = "0"] = match;
  const result = new Date(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), Number(second));
  if (
    Number.isNaN(result.getTime())
    || result.getFullYear() !== Number(year)
    || result.getMonth() !== Number(month) - 1
    || result.getDate() !== Number(day)
    || result.getHours() !== Number(hour)
    || result.getMinutes() !== Number(minute)
    || result.getSeconds() !== Number(second)
  ) return null;
  return result;
}

async function imageDate(file) {
  try {
    if (window.exifr) {
      const exif = await window.exifr.parse(file, { pick: ["DateTimeOriginal", "CreateDate", "ModifyDate"] });
      const value = exif && (exif.DateTimeOriginal || exif.CreateDate || exif.ModifyDate);
      if (value) {
        const result = value instanceof Date ? value : new Date(value);
        if (!Number.isNaN(result.getTime())) return { date: result, source: "EXIF拍摄时间" };
      }
    }
  } catch (_) {
    // Continue with filename or modified time when EXIF is unavailable.
  }
  const fromName = filenameDate(file);
  if (fromName) return { date: fromName, source: "文件名时间" };
  return { date: new Date(file.lastModified || Date.now()), source: "文件修改时间" };
}

async function collectImages(files) {
  const items = await Promise.all([...files].map(async file => {
    const time = await imageDate(file);
    return { file, date: time.date, source: time.source, group: groupName(time.date.getHours()) };
  }));
  return items.sort((left, right) => left.date - right.date);
}

function parseXml(text) {
  return new DOMParser().parseFromString(text, "application/xml");
}

function serializeXml(document) {
  return new XMLSerializer().serializeToString(document);
}

function coordinateKey(coordinate) {
  const match = coordinate.match(/^([A-Z]+)(\d+)$/i);
  if (!match) return [0, 0];
  let column = 0;
  for (const character of match[1].toUpperCase()) column = column * 26 + character.charCodeAt(0) - 64;
  return [Number(match[2]), column];
}

function targetCoordinates(anchor, count, columns = 2) {
  const match = anchor.match(/^([A-Z]+)(\d+)$/i);
  if (!match) throw new Error(`模板起始位置不正确：${anchor}`);
  let startColumn = 0;
  for (const character of match[1].toUpperCase()) startColumn = startColumn * 26 + character.charCodeAt(0) - 64;
  const startRow = Number(match[2]);
  const result = [];
  for (let index = 0; index < count; index += 1) {
    let column = startColumn + index % columns;
    let columnName = "";
    while (column) {
      const remainder = (column - 1) % 26;
      columnName = String.fromCharCode(65 + remainder) + columnName;
      column = Math.floor((column - 1) / 26);
    }
    result.push(`${columnName}${startRow + Math.floor(index / columns) * 2}`);
  }
  return result;
}

function nextRelationshipId(relationshipsDocument) {
  const ids = [...relationshipsDocument.documentElement.children]
    .map(node => Number((node.getAttribute("Id") || "").replace("rId", "")))
    .filter(Number.isFinite);
  return `rId${Math.max(0, ...ids) + 1}`;
}

function nextMediaName(zip) {
  const names = Object.keys(zip.files).map(name => {
    const match = name.match(/^xl\/media\/image(\d+)\.(?:png|jpe?g)$/i);
    return match ? Number(match[1]) : 0;
  });
  return `xl/media/image${Math.max(0, ...names) + 1}.jpg`;
}

function formulaText(cell) {
  return cell.getElementsByTagNameNS(MAIN_NS, "f")[0]?.textContent || "";
}

function imageIdFromFormula(formula) {
  return formula.match(/DISPIMG\(\\?"([^"\\]+)/)?.[1] || "";
}

function uniqueImageId() {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID().replace(/-/g, "").toUpperCase();
  }
  if (typeof crypto !== "undefined" && typeof crypto.getRandomValues === "function") {
    const bytes = crypto.getRandomValues(new Uint8Array(16));
    return [...bytes].map(value => value.toString(16).padStart(2, "0")).join("").toUpperCase();
  }
  return `${Date.now().toString(16)}${Math.random().toString(16).slice(2)}`.toUpperCase();
}

function addFormulaCell(sheetDocument, coordinate, imageId) {
  const sheetData = sheetDocument.getElementsByTagNameNS(MAIN_NS, "sheetData")[0];
  const rowNumber = coordinate.match(/\d+$/)[0];
  const row = [...sheetData.getElementsByTagNameNS(MAIN_NS, "row")].find(node => node.getAttribute("r") === rowNumber);
  if (!row) throw new Error(`模板没有图片行 ${rowNumber}，无法自动增加图片位 ${coordinate}。`);
  const cell = sheetDocument.createElementNS(MAIN_NS, "c");
  cell.setAttribute("r", coordinate);
  cell.setAttribute("t", "str");
  const formula = sheetDocument.createElementNS(MAIN_NS, "f");
  formula.textContent = `_xlfn.DISPIMG("${imageId}",1)`;
  const value = sheetDocument.createElementNS(MAIN_NS, "v");
  value.textContent = `=DISPIMG("${imageId}",1)`;
  cell.append(formula, value);
  row.appendChild(cell);
  return cell;
}

function addWpsSlot(sheetDocument, cellImagesDocument, relationshipsDocument, zip, templateItem, nextPictureId) {
  const imageId = `ID_AUTO_${uniqueImageId()}`;
  const relationshipId = nextRelationshipId(relationshipsDocument);
  const mediaName = nextMediaName(zip);
  const newItem = templateItem.cloneNode(true);
  const nameNode = newItem.getElementsByTagNameNS(DRAWING_NS, "cNvPr")[0];
  const blipNode = newItem.getElementsByTagNameNS(DRAWING_MAIN_NS, "blip")[0];
  nameNode.setAttribute("id", String(nextPictureId));
  nameNode.setAttribute("name", imageId);
  nameNode.setAttribute("descr", new Date().toISOString());
  blipNode.setAttributeNS(REL_NS, "r:embed", relationshipId);
  cellImagesDocument.documentElement.appendChild(newItem);
  const relationship = relationshipsDocument.createElementNS(REL_PACKAGE_NS, "Relationship");
  relationship.setAttribute("Id", relationshipId);
  relationship.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
  relationship.setAttribute("Target", `media/${mediaName.split("/").pop()}`);
  relationshipsDocument.documentElement.appendChild(relationship);
  return { imageId, relationshipId, mediaName, nextPictureId: nextPictureId + 1 };
}

async function loadDrawableImage(file) {
  if (typeof createImageBitmap === "function") {
    try {
      const bitmap = await createImageBitmap(file);
      return { image: bitmap, release: () => bitmap.close() };
    } catch (_) {
      // Older mobile browsers can fall back to a regular image element.
    }
  }
  return new Promise((resolve, reject) => {
    const sourceUrl = URL.createObjectURL(file);
    const image = new Image();
    image.onload = () => resolve({ image, release: () => URL.revokeObjectURL(sourceUrl) });
    image.onerror = () => {
      URL.revokeObjectURL(sourceUrl);
      reject(new Error(`无法读取图片：${file.name}`));
    };
    image.src = sourceUrl;
  });
}

function centerCrop(sourceWidth, sourceHeight) {
  const sourceAspect = sourceWidth / sourceHeight;
  const targetAspect = OUTPUT_IMAGE_WIDTH / OUTPUT_IMAGE_HEIGHT;
  let sourceX = 0;
  let sourceY = 0;
  let cropWidth = sourceWidth;
  let cropHeight = sourceHeight;
  if (sourceAspect > targetAspect) {
    cropWidth = sourceHeight * targetAspect;
    sourceX = (sourceWidth - cropWidth) / 2;
  } else if (sourceAspect < targetAspect) {
    cropHeight = sourceWidth / targetAspect;
    sourceY = (sourceHeight - cropHeight) / 2;
  }
  return { sourceX, sourceY, cropWidth, cropHeight };
}

async function imageAsJpeg(file) {
  const drawable = await loadDrawableImage(file);
  const canvas = document.createElement("canvas");
  canvas.width = OUTPUT_IMAGE_WIDTH;
  canvas.height = OUTPUT_IMAGE_HEIGHT;
  const drawingContext = canvas.getContext("2d");
  if (!drawingContext) {
    drawable.release();
    throw new Error("当前浏览器无法处理图片，请换用最新版浏览器。");
  }
  const sourceWidth = drawable.image.naturalWidth || drawable.image.width;
  const sourceHeight = drawable.image.naturalHeight || drawable.image.height;
  const { sourceX, sourceY, cropWidth, cropHeight } = centerCrop(sourceWidth, sourceHeight);
  drawingContext.drawImage(
    drawable.image,
    sourceX,
    sourceY,
    cropWidth,
    cropHeight,
    0,
    0,
    OUTPUT_IMAGE_WIDTH,
    OUTPUT_IMAGE_HEIGHT,
  );
  drawable.release();
  const blob = await new Promise(resolve => canvas.toBlob(resolve, "image/jpeg", 0.88));
  if (!blob) throw new Error(`无法读取图片：${file.name}`);
  return new Uint8Array(await blob.arrayBuffer());
}

async function replaceTemplate(files, secondRowText, onProgress = () => {}) {
  onProgress("正在下载内置模板……");
  const templateResponse = await fetch("template.xlsx");
  if (!templateResponse.ok) throw new Error("网页没有找到 template.xlsx，请确认它和网页在同一目录。");
  const templateBuffer = await templateResponse.arrayBuffer();
  onProgress("正在读取模板……");
  const zip = await JSZip.loadAsync(templateBuffer);
  const sheetDocument = parseXml(await zip.file("xl/worksheets/sheet1.xml").async("text"));
  const cellImagesDocument = parseXml(await zip.file("xl/cellimages.xml").async("text"));
  const relationshipsDocument = parseXml(await zip.file("xl/_rels/cellimages.xml.rels").async("text"));
  const stringsDocument = parseXml(await zip.file("xl/sharedStrings.xml").async("text"));
  const cellImages = [...cellImagesDocument.getElementsByTagNameNS(WPS_NS, "cellImage")];
  if (!cellImages.length) throw new Error("模板没有找到 WPS 图片位置。");

  const relationByImageId = new Map();
  for (const item of cellImages) {
    const nameNode = item.getElementsByTagNameNS(DRAWING_NS, "cNvPr")[0];
    const blipNode = item.getElementsByTagNameNS(DRAWING_MAIN_NS, "blip")[0];
    relationByImageId.set(nameNode.getAttribute("name"), blipNode.getAttributeNS(REL_NS, "embed"));
  }
  const relationTargets = new Map([...relationshipsDocument.documentElement.children].map(node => [node.getAttribute("Id"), node.getAttribute("Target")]));
  const formulaCells = new Map([...sheetDocument.getElementsByTagNameNS(MAIN_NS, "c")]
    .filter(cell => formulaText(cell).includes("DISPIMG"))
    .map(cell => [cell.getAttribute("r"), cell]));

  onProgress("正在识别照片时间并分组……");
  const allImages = await collectImages(files);
  const byGroup = Object.fromEntries(["上午", "中午", "下午", "未分类"].map(label => [label, allImages.filter(item => item.group === label)]));
  const selected = { 上午: [], 中午: [], 下午: [] };
  const used = new Set();
  const overflow = [];
  for (const label of GROUP_ORDER) {
    selected[label] = byGroup[label].slice(0, 2);
    selected[label].forEach(item => used.add(item.file));
    overflow.push(...byGroup[label].slice(2));
  }
  overflow.push(...byGroup["未分类"]);
  for (const label of GROUP_ORDER) {
    while (selected[label].length < 2 && overflow.length) {
      const candidate = overflow.shift();
      if (!used.has(candidate.file)) {
        selected[label].push(candidate);
        used.add(candidate.file);
      }
    }
  }

  let nextPictureId = Math.max(0, ...cellImages.map(item => Number(item.getElementsByTagNameNS(DRAWING_NS, "cNvPr")[0].getAttribute("id")) || 0)) + 1;
  const templateItem = cellImages[0];
  const slotImageIds = new Map([...formulaCells.entries()].map(([coordinate, cell]) => [coordinate, imageIdFromFormula(formulaText(cell))]));
  for (const label of GROUP_ORDER) {
    for (const coordinate of targetCoordinates(anchors[label], selected[label].length, 2)) {
      if (formulaCells.has(coordinate)) continue;
      const slot = addWpsSlot(sheetDocument, cellImagesDocument, relationshipsDocument, zip, templateItem, nextPictureId);
      nextPictureId = slot.nextPictureId;
      const newCell = addFormulaCell(sheetDocument, coordinate, slot.imageId);
      formulaCells.set(coordinate, newCell);
      slotImageIds.set(coordinate, slot.imageId);
      relationByImageId.set(slot.imageId, slot.relationshipId);
      relationTargets.set(slot.relationshipId, slot.mediaName);
    }
  }

  const assignedCoordinates = new Set(GROUP_ORDER.flatMap(label => targetCoordinates(anchors[label], selected[label].length, 2)));
  for (const [coordinate, cell] of formulaCells) {
    if (!assignedCoordinates.has(coordinate)) {
      const formula = cell.getElementsByTagNameNS(MAIN_NS, "f")[0];
      const value = cell.getElementsByTagNameNS(MAIN_NS, "v")[0];
      if (formula) formula.textContent = "";
      if (value) value.textContent = "";
    }
  }

  const usedCount = GROUP_ORDER.reduce((total, label) => total + selected[label].length, 0);
  let processedCount = 0;
  for (const label of GROUP_ORDER) {
    const coordinates = targetCoordinates(anchors[label], selected[label].length, 2);
    for (let index = 0; index < coordinates.length; index += 1) {
      const coordinate = coordinates[index];
      const item = selected[label][index];
      const cell = formulaCells.get(coordinate);
      const imageId = slotImageIds.get(coordinate);
      cell.getElementsByTagNameNS(MAIN_NS, "f")[0].textContent = `_xlfn.DISPIMG("${imageId}",1)`;
      cell.getElementsByTagNameNS(MAIN_NS, "v")[0].textContent = `=DISPIMG("${imageId}",1)`;
      const target = relationTargets.get(relationByImageId.get(imageId));
      const mediaName = target.startsWith("xl/") ? target : `xl/${target.replace(/^\//, "")}`;
      onProgress(`正在处理照片 ${processedCount + 1}/${usedCount}……`);
      zip.file(mediaName, await imageAsJpeg(item.file), { compression: "STORE" });
      processedCount += 1;
    }
  }

  const strings = [...stringsDocument.getElementsByTagNameNS(MAIN_NS, "si")];
  const date = new Date();
  const dateLabel = `${date.getMonth() + 1}月${date.getDate()}日`;
  const secondNode = strings[1]?.getElementsByTagNameNS(MAIN_NS, "t")[0];
  if (secondNode) secondNode.textContent = secondRowText.trim() ? secondRowText.replace(/\{\{日期\}\}/g, dateLabel) : (secondNode.textContent.replace(/^\s*\d{1,2}月\d{1,2}日/, dateLabel) || dateLabel);

  zip.file("xl/worksheets/sheet1.xml", serializeXml(sheetDocument));
  zip.file("xl/cellimages.xml", serializeXml(cellImagesDocument));
  zip.file("xl/_rels/cellimages.xml.rels", serializeXml(relationshipsDocument));
  zip.file("xl/sharedStrings.xml", serializeXml(stringsDocument));
  onProgress("正在打包 Excel……");
  const blob = await zip.generateAsync({
    type: "blob",
    compression: "DEFLATE",
    compressionOptions: { level: 1 },
    streamFiles: true,
  });
  return { blob, usedCount, ignoredCount: Math.max(0, allImages.length - usedCount) };
}

function download(blob, filename) {
  if (currentDownloadUrl) URL.revokeObjectURL(currentDownloadUrl);
  const url = URL.createObjectURL(blob);
  currentDownloadUrl = url;
  downloadFallback.href = url;
  downloadFallback.download = filename;
  downloadFallback.hidden = false;
  downloadFallback.click();
}

generateButton.addEventListener("click", async () => {
  if (!photosInput.files.length) {
    setStatus("请先选择施工照片。", "error");
    return;
  }
  generateButton.disabled = true;
  downloadFallback.hidden = true;
  setStatus("正在读取时间、匹配分组并生成 Excel……");
  try {
    const result = await replaceTemplate(photosInput.files, secondRowInput.value, message => setStatus(message));
    const date = new Date();
    download(result.blob, `海滨大道施工日志${date.getMonth() + 1}.${date.getDate()}_已生成.xlsx`);
    const ignoredText = result.ignoredCount ? `，另有 ${result.ignoredCount} 张未使用` : "";
    setStatus(`生成完成，已使用 ${result.usedCount} 张照片${ignoredText}，Excel 已开始下载；如未自动下载，请点击下方链接。`, "success");
  } catch (error) {
    console.error(error);
    setStatus(`生成失败：${error.message || error}`, "error");
  } finally {
    generateButton.disabled = false;
  }
});
