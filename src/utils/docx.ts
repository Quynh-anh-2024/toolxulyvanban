export interface FormattingConfig {
  font: string;
  size: number;
  spacing: string;
  textAlign: string;
  textIndent: number;
  paraSpacing: number;
  leftMargin: number;
  rightMargin: number;
}

export function exportToWord(
  previewElement: HTMLElement | null,
  fileName: string,
  config: FormattingConfig
) {
  if (!previewElement) return;

  const {
    font,
    size,
    spacing,
    textAlign,
    textIndent,
    paraSpacing,
    leftMargin,
    rightMargin,
  } = config;

  let lineSpacingPercent = "150%";
  if (spacing === "1.0") lineSpacingPercent = "100%";
  if (spacing === "1.15") lineSpacingPercent = "115%";
  if (spacing === "1.5") lineSpacingPercent = "150%";

  // Clone the element to remove injected React styles or classes if needed, 
  // but we mostly just need its inner HTML. We strip out unnecessary classes.
  const cloneDiv = previewElement.cloneNode(true) as HTMLElement;
  cloneDiv.querySelectorAll("*").forEach((el) => {
    // Preserve inline styles for images to retain dimensions from original file
    if (el.tagName.toLowerCase() !== "img") {
      el.removeAttribute("style");
    }
  });
  
  // Extract base64 images and substitute with Content-IDs
  const imagesToEmbed: { id: string; type: string; data: string }[] = [];
  
  cloneDiv.querySelectorAll("img").forEach((img, index) => {
    const src = img.getAttribute("src");
    if (src && src.startsWith("data:image/")) {
      const parts = src.match(/^data:(image\/[^;]+);base64,(.*)$/);
      if (parts && parts.length === 3) {
        const id = `image_${index}`;
        imagesToEmbed.push({
          id: id,
          type: parts[1],
          data: parts[2],
        });
        img.setAttribute("src", `cid:${id}`);
      }
    }
  });

  const cleanHTML = cloneDiv.innerHTML;

  const fullHTML = `
  <html xmlns:o="urn:schemas-microsoft-com:office:office" 
        xmlns:w="urn:schemas-microsoft-com:office:word" 
        xmlns="http://www.w3.org/TR/REC-html40">
  <head>
      <meta charset="utf-8">
      <title>Export</title>
      <style>
          body {
              margin: 0; padding: 0;
              font-family: "${font}", serif;
              font-size: ${size}pt;
          }
          @page {
              mso-page-orientation: portrait;
              size: 21cm 29.7cm;
              margin: 2cm ${rightMargin}cm 2cm ${leftMargin}cm;
          }
          @page WordSection1 {
              size: 21cm 29.7cm;
              margin: 2cm ${rightMargin}cm 2cm ${leftMargin}cm;
              mso-header-margin: 35.4pt;
              mso-footer-margin: 35.4pt;
              mso-paper-source: 0;
          }
          div.WordSection1 { page: WordSection1; }
          
          p {
              margin: 0; padding: 0;
              font-family: "${font}", serif;
              font-size: ${size}pt;
              text-align: ${textAlign};
              text-indent: ${textIndent}cm;
              margin-bottom: ${paraSpacing}pt;
              line-height: ${lineSpacingPercent}; 
              mso-line-height-rule: exactly;
              
              mso-pagination: none;
              page-break-inside: auto;
          }

          /* Reset paragraph styles inside tables */
          table p {
              text-align: left !important;
              text-indent: 0cm !important;
              margin-bottom: 0pt !important;
              line-height: normal !important;
          }
          th, th p {
              text-align: center !important;
              font-weight: bold !important;
          }
          /* Cột STT căn giữa */
          tr td:first-child, tr td:first-child p {
              text-align: center !important;
          }
          tr td:not(:first-child), tr td:not(:first-child) p {
              text-align: left !important;
          }

          /* MS WORD CSS FIXES for AI blocks */
          div.ai-wrapper-box {
              border-top: 1pt dashed #000;
              margin-top: 20pt;
              padding-top: 10pt;
          }
          div.ai-wrapper-box p.ai-title-line {
              text-align: center !important;
              text-indent: 0cm !important;
              font-size: 14pt;
              font-weight: bold;
              margin-bottom: 12pt !important;
          }
          div.ai-wrapper-box p.ai-text-line {
              text-align: left !important;
              text-indent: 0cm !important;
              margin-top: 4pt !important;
              margin-bottom: 4pt !important;
          }

          h1, h2, h3, h4, h5, h6 {
              font-family: "${font}", serif;
              mso-pagination: none;
              page-break-inside: auto;
              page-break-after: auto;
          }

          table { border-collapse: collapse; width: 100%; margin-bottom: 12pt; }
          td, th { border: 1px solid black; padding: 4pt; }
          
          /* Do not force image size, let it use its natural or parsed style dimensions */
          img { display: inline-block; }
      </style>
  </head>
  <body>
      <div class="WordSection1">
          ${cleanHTML}
      </div>
  </body>
  </html>
  `;

  // Construct MHTML
  const boundary = "----=_NextPart_MHTML_BOUNDARY";
  
  // Convert HTML to base64 safely (handling UTF-8)
  const htmlBase64 = btoa(unescape(encodeURIComponent(fullHTML)));

  let mhtml = `MIME-Version: 1.0\r\n`;
  mhtml += `Content-Type: multipart/related; boundary="${boundary}"\r\n\r\n`;
  
  mhtml += `--${boundary}\r\n`;
  mhtml += `Content-Type: text/html; charset="utf-8"\r\n`;
  mhtml += `Content-Transfer-Encoding: base64\r\n\r\n`;
  mhtml += `${htmlBase64.replace(/(.{76})/g, "$1\r\n")}\r\n\r\n`;

  for (const img of imagesToEmbed) {
    mhtml += `--${boundary}\r\n`;
    mhtml += `Content-Type: ${img.type}\r\n`;
    mhtml += `Content-Transfer-Encoding: base64\r\n`;
    mhtml += `Content-ID: <${img.id}>\r\n\r\n`;
    mhtml += `${img.data.replace(/(.{76})/g, "$1\r\n")}\r\n\r\n`;
  }

  mhtml += `--${boundary}--\r\n`;

  const blob = new Blob([mhtml], { type: "application/msword" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `${fileName}_ChuanND30.doc`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(a.href);
}

export function processHtmlToRemoveBullets(htmlString: string): string {
  const tempDiv = document.createElement("div");
  tempDiv.innerHTML = htmlString;

  const listItems = tempDiv.querySelectorAll("li");
  listItems.forEach((li) => {
    const p = document.createElement("p");
    p.innerHTML = li.innerHTML;
    if (li.parentNode) {
      li.parentNode.replaceChild(p, li);
    }
  });

  const lists = tempDiv.querySelectorAll("ul, ol");
  lists.forEach((list) => {
    const fragment = document.createDocumentFragment();
    while (list.firstChild) {
      fragment.appendChild(list.firstChild);
    }
    if (list.parentNode) {
      list.parentNode.replaceChild(fragment, list);
    }
  });

  return tempDiv.innerHTML;
}
