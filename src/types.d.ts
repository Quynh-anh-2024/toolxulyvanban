declare module 'mammoth' {
  export interface ConvertToHtmlOptions {
    arrayBuffer?: ArrayBuffer;
    path?: string;
  }

  export interface ExtractRawTextOptions {
    arrayBuffer?: ArrayBuffer;
    path?: string;
  }

  export interface ConversionResult {
    value: string;
    messages: any[];
  }

  export function convertToHtml(
    input: ConvertToHtmlOptions,
    options?: any
  ): Promise<ConversionResult>;

  export function extractRawText(
    input: ExtractRawTextOptions
  ): Promise<ConversionResult>;

  export const images: any;
}

interface ImportMetaEnv {
  readonly VITE_OPENROUTER_API_KEY?: string;
  readonly VITE_GEMINI_API_KEY?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
