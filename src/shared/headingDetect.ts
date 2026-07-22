/** 标题/题注启发式识别（扫描与套用共用） */

export function detectHeadingLevel(text: string): 1 | 2 | 3 | null {
  const t = text.replace(/\s+/g, " ").trim();
  if (!t || t.length > 60) return null;

  if (/^第[一二三四五六七八九十百零〇两\d]+[章节]/.test(t)) return 1;
  if (/^(摘要|目录|参考文献|致谢|附录)/.test(t)) return 1;
  if (/^[一二三四五六七八九十百]+[、．.]\s*\S/.test(t)) return 1;
  if (/^[（(][一二三四五六七八九十\d]+[)）]/.test(t)) return 2;
  if (/^\d+([、．.]|\s+)/.test(t) && !/^\d+\.\d+/.test(t)) return 2;
  if (/^\d+\.\d+([^\d]|$)/.test(t) && !/^\d+\.\d+\.\d+/.test(t)) return 2;
  if (/^\d+\.\d+\.\d+/.test(t)) return 3;
  return null;
}

export function detectCaption(text: string): boolean {
  const t = text.replace(/\s+/g, " ").trim();
  return /^(图|表|Figure|Table|Fig\.?)\s*[\d一二三四五六七八九十]+/i.test(t);
}
