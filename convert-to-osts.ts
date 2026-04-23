import { readdirSync, readFileSync, statSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

const SCRIPTS_DIR = resolve(import.meta.dirname, 'scripts');

interface Param {
  name: string;
  type: string;
}

interface JsonSchemaType {
  type: string;
  items?: JsonSchemaType;
}

interface OstsFile {
  version: string;
  body: string;
  description: string;
  noCodeMetadata: string;
  parameterInfo: string;
  apiInfo: string;
}

function tsTypeToJsonSchema(tsType: string): JsonSchemaType {
  const type = tsType.trim();

  if (type.endsWith('[]')) {
    return { type: 'array', items: tsTypeToJsonSchema(type.slice(0, -2)) };
  }

  switch (type) {
    case 'string':  return { type: 'string' };
    case 'number':  return { type: 'number' };
    case 'boolean': return { type: 'boolean' };
    default:        return { type: 'string' };
  }
}

function parseMainParams(code: string): Param[] {
  const match = code.match(/function\s+main\s*\(([\s\S]*?)\)\s*[:{]/);
  if (!match) throw new Error('No `main` function found.');

  const paramsStr = match[1];
  const rawParams: string[] = [];
  let depth = 0;
  let current = '';

  for (const ch of paramsStr) {
    if ('<(['.includes(ch)) depth++;
    else if ('>)]'.includes(ch)) depth--;
    else if (ch === ',' && depth === 0) {
      const trimmed = current.trim();
      if (trimmed) rawParams.push(trimmed);
      current = '';
      continue;
    }
    current += ch;
  }
  const last = current.trim();
  if (last) rawParams.push(last);

  return rawParams.map(param => {
    const normalised = param.replace(/\s+/g, ' ').trim();
    const colonIdx = normalised.indexOf(':');
    if (colonIdx === -1) return { name: normalised, type: 'string' };
    return {
      name: normalised.slice(0, colonIdx).trim(),
      type: normalised.slice(colonIdx + 1).trim(),
    };
  });
}

function parseDescription(code: string): string {
  const jsdocMatch = code.match(/\/\*\*([\s\S]*?)\*\/\s*function\s+main/);
  if (!jsdocMatch) return '';

  const descLines: string[] = [];
  for (const line of jsdocMatch[1].split('\n')) {
    const cleaned = line.replace(/^\s*\*\s?/, '').trim();
    if (cleaned.startsWith('@')) break;
    if (cleaned) descLines.push(cleaned);
  }

  return descLines.join(' ');
}

function convertToOsts(inputPath: string): void {
  const code = readFileSync(inputPath, 'utf8');
  const body = code.replace(/\r\n/g, '\n').replace(/\n/g, '\r\n');

  const description = parseDescription(code);
  const allParams = parseMainParams(code);
  const userParams = allParams.filter(p => p.name !== 'workbook');

  const properties: Record<string, JsonSchemaType> = {};
  for (const p of userParams) {
    properties[p.name] = tsTypeToJsonSchema(p.type);
  }

  const parameterInfo = {
    version: 1,
    originalParameterOrder: userParams.map((p, i) => ({ name: p.name, index: i })),
    parameterSchema:
      userParams.length > 0
        ? { type: 'object', required: userParams.map(p => p.name), properties }
        : { type: 'object', properties: {} },
    returnSchema: { type: 'object', properties: {} },
    signature: {
      comment: '',
      parameters: allParams.map(p => ({ name: p.name, comment: '' })),
    },
  };

  const osts: OstsFile = {
    version: '0.3.0',
    body,
    description,
    noCodeMetadata: '',
    parameterInfo: JSON.stringify(parameterInfo),
    apiInfo: JSON.stringify({ variant: 'synchronous', variantVersion: 2 }),
  };

  const outputPath = inputPath.replace(/\.ts$/, '.osts');
  writeFileSync(outputPath, JSON.stringify(osts));
  console.log(`Converted: ${outputPath}`);
}

function walkAndConvert(dir: string): void {
  for (const entry of readdirSync(dir)) {
    const fullPath = join(dir, entry);
    if (statSync(fullPath).isDirectory()) {
      walkAndConvert(fullPath);
    } else if (entry.endsWith('.ts')) {
      try {
        convertToOsts(fullPath);
      } catch (err) {
        console.error(`Failed: ${fullPath} — ${(err as Error).message}`);
        process.exitCode = 1;
      }
    }
  }
}

walkAndConvert(SCRIPTS_DIR);
