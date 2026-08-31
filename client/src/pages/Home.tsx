/* Diseño Cartografía Cívica: consulta multindicador con jerarquía institucional y contexto territorial. */
/** Estilo Cartografía Cívica: consulta territorial clara, jerarquía institucional y resultados trazables. */
import { type CSSProperties, useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import embeddedJson from "@/data/indicadores.json?raw";
import { feature as topoFeature } from "topojson-client";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { BarChart3, BookOpen, ChevronDown, ChevronUp, CircleHelp, Contrast, Download, FileSpreadsheet, Filter, Info, LineChart as LineChartIcon, Loader2, MapPin, Minus, Plus, Search, SlidersHorizontal, X } from "lucide-react";
import { clearDimensionSelections } from "@/lib/filterState";
import { Bar, CartesianGrid, ComposedChart, LabelList, Line, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";

type Value = string | number | null | undefined;
type Row = Record<string, Value>;
type Filters = Record<string, string>;
type FilterConfig = { key: string; label: string };
type MunicipalFeature = { properties: { dpt?: string; name?: string }; geometry: { type: string; coordinates: unknown } };
type Series = { name: string; points: Map<number, number> };
type Variation = { name: string; items: { label: string; difference: number | null; relative: number | null }[] };
type WorkbookSheetSummary = { name: string; rowCount: number; fieldCount: number };
type TechnicalProfile = { code: string; fields: Record<string, string>; markedDimensions: string[] };
type EmbeddedData = { rows: Row[]; technicalProfiles: TechnicalProfile[]; workbookStructure: WorkbookSheetSummary[] };
const EMBEDDED_DATA = JSON.parse(embeddedJson) as EmbeddedData;

const LOGO_URL = "./assets/logo_gobernacion_putumayo.webp";
const SYMBOL_URL = "./assets/putumayo-data-symbol.png";
const LANDSCAPE_URL = "./assets/putumayo-data-landscape.jpg";
const TEXTURE_URL = "./assets/putumayo-microtexture.jpg";
const MUNICIPAL_MAP_URL = "./assets/colombia-municipios.topojson";
const MAIN_SHEET = "Base Indicadores";
const ALL = "__TODOS__";
const SPECIAL_KEY = "DESAGREGACION ESPECIAL";
const DEPARTMENT_KEY = "TERRITORIAL DEPARTAMENTO";
const MUNICIPALITY_KEY = "TERRITORIAL MUNICIPIO";
const SPECIAL_DIRECT_CODES = new Set(["IND-DANE-016", "IND-DANE-026", "IND-DANE-034", "IND-DANE-035", "IND-DANE-074"]);
const DIRECT_TERRITORIAL_CODES = new Set(["IND-DANE-032", "IND-DANE-050", "IND-DANE-051", "IND-DANE-054", "IND-DANE-056", "IND-DANE-057", "IND-DANE-064"]);
const VARIANT_REQUIRED_CODES = new Set(["IND-DANE-016", "IND-DANE-034", "IND-DANE-035"]);
const PRINCIPAL_DIRECT_CODES = new Set(["IND-DANE-026", "IND-DANE-074"]);
const DIRECT_DIMENSION_CODE = "IND-DANE-032";

const primary: FilterConfig[] = [
  { key: "AREA TEMATICA", label: "Área temática" },
  { key: "TEMA", label: "Tema" },
  { key: "INDICADOR", label: "Indicador" },
  { key: "VIGENCIA", label: "Vigencia" },
];
const dimensions: FilterConfig[] = [
  { key: DEPARTMENT_KEY, label: "Departamento" },
  { key: SPECIAL_KEY, label: "Desagregación especial" },
  { key: MUNICIPALITY_KEY, label: "Municipio" },
  { key: "ÁREA_GEOGRÁFICA", label: "Área geográfica" },
  { key: "SEXO", label: "Sexo" },
  { key: "NIVEL_EDUCATIVO_GENERAL", label: "Nivel educativo general" },
  { key: "NIVEL_EDUCATIVO SUPERIOR", label: "Especialista" },
];
const variationDimensionKeys = new Set([DEPARTMENT_KEY, ...dimensions.map((field) => field.key)]);
const dataOrder = [
  "ID_INDICADOR", "FUENTE_ORIGEN", "TEMA", "SUBTEMA", "INDICADOR", "DESCRIPCION", "FUENTE", "LINK FUENTE",
  "PERIODICIDAD_REPORTE", "UNIDAD_MEDIDA", "VIGENCIA", "RESULTADO", "CODIDO DIVIPOLA", "ESTADO", DEPARTMENT_KEY, MUNICIPALITY_KEY,
  "ÁREA_GEOGRÁFICA", "SEXO", "NIVEL_EDUCATIVO_GENERAL", "NIVEL_EDUCATIVO SUPERIOR", "OBJETIVO", "META", "FECHA ULTIMA CONSULTA",
];
const labels: Record<string, string> = {
  ID_INDICADOR: "Código", FUENTE_ORIGEN: "Fuente de origen", TEMA: "Tema", SUBTEMA: "Subtema", INDICADOR: "Indicador",
  DESCRIPCION: "Descripción", FUENTE: "Fuente", "LINK FUENTE": "Enlace fuente", PERIODICIDAD_REPORTE: "Periodicidad",
  UNIDAD_MEDIDA: "Unidad", VIGENCIA: "Vigencia", RESULTADO: "Resultado", "CODIDO DIVIPOLA": "Código DIVIPOLA",
  ESTADO: "Estado", [DEPARTMENT_KEY]: "Departamento", [MUNICIPALITY_KEY]: "Municipio", "ÁREA_GEOGRÁFICA": "Área geográfica", SEXO: "Sexo", GENERO: "Género",
  GRUPO_ETNICO: "Grupo étnico", DISCAPACIDAD: "Discapacidad", ETAPAS_DEL_CURSO_VIDA: "Etapas del curso de vida",
  NIVEL_EDUCATIVO_GENERAL: "Nivel educativo general", "NIVEL_EDUCATIVO SUPERIOR": "Nivel educativo superior",
  OBJETIVO: "ODS", META: "Meta ODS", "FECHA ULTIMA CONSULTA": "Fecha de última consulta", "AREA TEMATICA": "Área temática",
  [SPECIAL_KEY]: "Desagregación especial",
};
const metadataFieldGroups = [
  { title: "Identificación y clasificación", fields: ["ID_INDICADOR", "FUENTE_ORIGEN", "TEMA", "SUBTEMA", "INDICADOR", "DESCRIPCION"] },
  { title: "Medición y temporalidad", fields: ["PERIODICIDAD_REPORTE", "UNIDAD_MEDIDA", "VIGENCIA", "RESULTADO", "ESTADO", "FECHA ULTIMA CONSULTA"] },
  { title: "Cobertura y desagregación", fields: ["CODIDO DIVIPOLA", DEPARTMENT_KEY, MUNICIPALITY_KEY, "ÁREA_GEOGRÁFICA", "SEXO", "NIVEL_EDUCATIVO_GENERAL", "NIVEL_EDUCATIVO SUPERIOR"] },
  { title: "Fuentes y Agenda 2030", fields: ["FUENTE", "LINK FUENTE", "OBJETIVO", "META"] },
];
const metadataFieldNotes: Record<string, string> = {
  ID_INDICADOR: "Código institucional del indicador.", FUENTE_ORIGEN: "Entidad o sistema de procedencia.", TEMA: "Área temática de clasificación.", SUBTEMA: "Tema específico de clasificación.", INDICADOR: "Nombre oficial del indicador.", DESCRIPCION: "Definición o alcance del indicador.",
  PERIODICIDAD_REPORTE: "Frecuencia de actualización reportada.", UNIDAD_MEDIDA: "Unidad en que se expresa el resultado.", VIGENCIA: "Año o periodo del dato.", RESULTADO: "Valor oficial reportado.", ESTADO: "Condición de publicación del dato.", "FECHA ULTIMA CONSULTA": "Fecha de referencia reportada en la fuente.",
  "CODIDO DIVIPOLA": "Código territorial DIVIPOLA disponible en la base.", [DEPARTMENT_KEY]: "Filtro territorial principal e independiente; Putumayo identifica la fila departamental reportada.", [MUNICIPALITY_KEY]: "Municipio disponible como filtro territorial complementario cuando la base reporta desagregación municipal.", "ÁREA_GEOGRÁFICA": "Área geográfica disponible para desagregar.", SEXO: "Desagregación por sexo.", GENERO: "Desagregación por género.", GRUPO_ETNICO: "Desagregación por grupo étnico.", DISCAPACIDAD: "Desagregación por discapacidad.", ETAPAS_DEL_CURSO_VIDA: "Desagregación por etapa del curso de vida.", NIVEL_EDUCATIVO_GENERAL: "Desagregación por nivel educativo general.", "NIVEL_EDUCATIVO SUPERIOR": "Desagregación por nivel educativo superior.",
  FUENTE: "Entidad responsable o fuente de información.", "LINK FUENTE": "Enlace de consulta de la fuente.", OBJETIVO: "Objetivo de Desarrollo Sostenible relacionado.", META: "Meta ODS relacionada.",
};
const mapCallouts: Record<string, { dx: number; dy: number }> = {
  sibundoy: { dx: 0, dy: -12 }, colon: { dx: -68, dy: 16 }, "san francisco": { dx: -70, dy: -16 },
  santiago: { dx: -76, dy: 18 }, "valle del guamuez": { dx: -30, dy: 36 }, "san miguel": { dx: -36, dy: 52 },
};
const colors = ["#004884", "#168b6b", "#c48a21", "#5a6b81", "#875f96", "#b85d3a", "#1c6e8c", "#577541", "#ad5a76", "#6e708e"];

const text = (value: Value) => String(value ?? "").replace(/\s+/g, " ").trim();
const normalized = (value: Value) => text(value).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
const numeric = (value: Value) => { const raw = text(value); const normalizedValue = raw.includes(",") ? raw.replace(/\./g, "").replace(",", ".") : /^\d{1,3}(\.\d{3})+$/.test(raw) ? raw.replace(/\./g, "") : raw; const parsed = Number(normalizedValue); return Number.isFinite(parsed) ? parsed : null; };
const fmt = (value: Value) => { const parsed = numeric(value); return parsed === null ? "No disponible" : new Intl.NumberFormat("es-CO", { maximumFractionDigits: 2, useGrouping: false }).format(parsed); };
const formatDate = (value: Value) => {
  const raw = text(value);
  const match = raw.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (!match) return raw || "No disponible";
  const first = Number(match[1]); const second = Number(match[2]);
  const year = match[3].length === 2 ? 2000 + Number(match[3]) : Number(match[3]);
  const day = first > 12 ? first : second > 12 ? second : first;
  const month = first > 12 ? second : second > 12 ? first : second;
  return `${String(day).padStart(2, "0")}/${String(month).padStart(2, "0")}/${year}`;
};
const short = (value: Value, limit = 150) => { const content = text(value); return content.length > limit ? `${content.slice(0, limit)}…` : content; };
const isWebLink = (value: Value) => /^https?:\/\//i.test(text(value));
const rowValue = (row: Row | undefined, key: string) => text(row?.[key]);
const cleanRow = (row: Row): Row => Object.fromEntries(Object.entries(row).map(([key, value]) => [text(key).replace(/\s+/g, " ").toUpperCase(), typeof value === "string" ? text(value) : value]));
const distinct = (values: string[]) => Array.from(new Set(values.filter(Boolean)));
const hasMeaningfulValue = (value: Value) => Boolean(text(value)) && !["no disponible", "n/a", "na", "-"].includes(normalized(value));

function dimensionFromFicha(code: string, label: string) {
  const candidate = normalized(label);
  if (candidate.includes("municipal")) return MUNICIPALITY_KEY;
  if (candidate.includes("urbano") || candidate.includes("rural") || candidate === "area") return "ÁREA_GEOGRÁFICA";
  if (candidate.includes("sexo")) return "SEXO";
  if (code === DIRECT_DIMENSION_CODE && candidate.includes("escolaridad")) return "NIVEL_EDUCATIVO_GENERAL";
  if (candidate.includes("especialista") || candidate.includes("especializacion") || candidate.includes("nivel educativo superior")) return "NIVEL_EDUCATIVO SUPERIOR";
  if (SPECIAL_DIRECT_CODES.has(code) && !candidate.includes("departamental") && !candidate.includes("regional")) return SPECIAL_KEY;
  return "";
}

function parseTechnicalProfile(sheet: XLSX.WorkSheet): TechnicalProfile | null {
  const matrix = XLSX.utils.sheet_to_json<Value[]>(sheet, { header: 1, defval: "", raw: false }) as Value[][];
  const codeMatch = matrix.flat().map(text).map((value) => value.match(/IND-DANE-(\d{1,3})/i)?.[1]).find(Boolean);
  if (!codeMatch) return null;
  const code = `IND-DANE-${codeMatch.padStart(3, "0")}`;
  const fields: Record<string, string> = {};
  const markedDimensions = new Set<string>();
  matrix.forEach((row, rowIndex) => {
    const fieldCode = text(row[0]).toUpperCase();
    const label = text(row[1]);
    const value = row.slice(2).map(text).filter(Boolean).at(-1) ?? "";
    if (/^B\d+$/.test(fieldCode) && !["B12", "B16"].includes(fieldCode) && label && value) fields[label] = value;
    row.forEach((cell, columnIndex) => {
      if (normalized(cell) !== "x") return;
      const markedLabel = row.slice(columnIndex + 1).map(text).find(Boolean) ?? "";
      const key = dimensionFromFicha(code, markedLabel);
      if (key) markedDimensions.add(key);
    });
  });
  return { code, fields, markedDimensions: Array.from(markedDimensions) };
}

function valueFor(row: Row | undefined, key: string, specialIndex: Map<string, Map<string, string>>) {
  if (key === "INDICADOR") return rowValue(row, "ID_INDICADOR");
  if (key === "AREA TEMATICA") return rowValue(row, "AREA TEMATICA") || rowValue(row, "TEMA");
  if (key === "TEMA") return rowValue(row, "SUBTEMA") || rowValue(row, "TEMA");
  if (key === "OBJETIVO ODS") return rowValue(row, "OBJETIVO ODS") || rowValue(row, "OBJETIVO");
  if (key === "META ODS") return rowValue(row, "META ODS") || rowValue(row, "META");
  if (key === "CODIGO DIVIPOLA") return rowValue(row, "CODIGO DIVIPOLA") || rowValue(row, "CODIDO DIVIPOLA");
  if (key === DEPARTMENT_KEY) return rowValue(row, DEPARTMENT_KEY) || (rowValue(row, MUNICIPALITY_KEY) ? "" : "Putumayo");
  if (key === MUNICIPALITY_KEY) return rowValue(row, MUNICIPALITY_KEY);
  if (key === SPECIAL_KEY) return specialIndex.get(rowValue(row, "ID_INDICADOR"))?.get(rowValue(row, "INDICADOR")) ?? "";
  return rowValue(row, key);
}

function isDirectTerritorialRow(row: Row, specialIndex: Map<string, Map<string, string>>) {
  const hasStandardDisaggregation = dimensions.filter((field) => ![DEPARTMENT_KEY, MUNICIPALITY_KEY, SPECIAL_KEY].includes(field.key)).some((field) => Boolean(valueFor(row, field.key, specialIndex)));
  const special = normalized(valueFor(row, SPECIAL_KEY, specialIndex));
  return !hasStandardDisaggregation && (!special || special === "total");
}

function aggregateResult(rows: Row[], specialIndex: Map<string, Map<string, string>>, preferDirectTotal = false) {
  const code = rowValue(rows[0], "ID_INDICADOR");
  const directRows = rows.filter((row) => isDirectTerritorialRow(row, specialIndex));
  const numericRows = rows.filter((row) => numeric(row.RESULTADO) !== null);
  if (SPECIAL_DIRECT_CODES.has(code) || DIRECT_TERRITORIAL_CODES.has(code) || code === DIRECT_DIMENSION_CODE) { const row = numericRows[0] ?? rows[0]; return { row, value: numeric(row?.RESULTADO), aggregated: false, count: 1 }; }
  if (preferDirectTotal) {
    const departmentDirect = directRows.filter((row) => !valueFor(row, MUNICIPALITY_KEY, specialIndex));
    const firstMunicipality = valueFor(rows[0], MUNICIPALITY_KEY, specialIndex);
    const municipalityDirect = firstMunicipality ? directRows.filter((row) => normalized(valueFor(row, MUNICIPALITY_KEY, specialIndex)) === normalized(firstMunicipality)) : [];
    const candidateRows = firstMunicipality ? municipalityDirect : departmentDirect;
    if (candidateRows.length === 1) return { row: candidateRows[0], value: numeric(candidateRows[0].RESULTADO), aggregated: false, count: 1 };
  }
  return { row: numericRows[0] ?? rows[0], value: numericRows.reduce((total, row) => total + (numeric(row.RESULTADO) ?? 0), 0), aggregated: numericRows.length > 1, count: numericRows.length };
}

function aggregateSeries(rows: Row[], labelFor: (row: Row) => string, specialIndex: Map<string, Map<string, string>>, preferDirectTotal: boolean) {
  const grouped = new Map<string, Map<number, Row[]>>();
  rows.forEach((row) => { const label = labelFor(row); const year = numeric(row.VIGENCIA); if (year === null || numeric(row.RESULTADO) === null) return; if (!grouped.has(label)) grouped.set(label, new Map()); const points = grouped.get(label)!; points.set(year, [...(points.get(year) ?? []), row]); });
  return Array.from(grouped.entries()).map(([name, rowsByYear]) => ({ name, points: new Map(Array.from(rowsByYear.entries()).map(([year, items]) => [year, aggregateResult(items, specialIndex, preferDirectTotal).value ?? 0])) }));
}

function sharedStart(values: string[]) {
  if (!values.length) return "";
  let prefix = values[0];
  for (const value of values.slice(1)) while (prefix && !normalized(value).startsWith(normalized(prefix))) prefix = prefix.slice(0, -1);
  return prefix.trimEnd();
}

function createSpecialIndex(rows: Row[]) {
  const index = new Map<string, Map<string, string>>();
  for (const code of distinct(rows.map((row) => rowValue(row, "ID_INDICADOR")))) {
    const variants = distinct(rows.filter((row) => rowValue(row, "ID_INDICADOR") === code).map((row) => rowValue(row, "INDICADOR")));
    if (!SPECIAL_DIRECT_CODES.has(code) || variants.length < 2) continue;
    const prefix = sharedStart(variants).replace(/[\s\-–—:;(]+$/, "").trim();
    const items = new Map<string, string>();
    variants.forEach((variant) => {
      const suffix = variant.slice(prefix.length).replace(/^[\s\-–—:;()]+|[\s\-–—:;()]+$/g, "").trim() || "Total";
      const grade = ["IND-DANE-034", "IND-DANE-035"].includes(code) ? normalized(variant).match(/\bgrados?\s*(5|9)\b/)?.[1] : undefined;
      items.set(variant, grade ? `${grade} - Grado ${grade}` : suffix);
    });
    index.set(code, items);
  }
  return index;
}

function matchesFilters(rows: Row[], filters: Filters, specialIndex: Map<string, Map<string, string>>, ignored?: string) {
  return rows.filter((row) => Object.entries(filters).every(([key, value]) => {
    if (key === ignored || !value || value === ALL) return true;
    if (key === DEPARTMENT_KEY && normalized(value) === "putumayo") {
      return normalized(valueFor(row, DEPARTMENT_KEY, specialIndex)) === "putumayo" || Boolean(valueFor(row, MUNICIPALITY_KEY, specialIndex));
    }
    return normalized(valueFor(row, key, specialIndex)) === normalized(value);
  }));
}

function emptyFilters(): Filters { return { ...Object.fromEntries([...primary, ...dimensions].map((field) => [field.key, ALL])), [DEPARTMENT_KEY]: ALL }; }

const positionsOf = (coordinates: unknown, points: number[][] = []): number[][] => { if (Array.isArray(coordinates) && typeof coordinates[0] === "number") points.push(coordinates as number[]); else if (Array.isArray(coordinates)) coordinates.forEach((item) => positionsOf(item, points)); return points; };
const mapPath = (geometry: any, bounds: { minX: number; minY: number; scale: number } | null) => { if (!geometry || !bounds) return ""; const polygons = geometry.type === "Polygon" ? [geometry.coordinates] : geometry.type === "MultiPolygon" ? geometry.coordinates : []; return polygons.map((polygon: any) => polygon.map((ring: any) => ring.map(([lng, lat]: number[], index: number) => `${index ? "L" : "M"}${24 + (lng - bounds.minX) * bounds.scale} ${356 - (lat - bounds.minY) * bounds.scale}`).join(" ") + " Z").join(" ")).join(" "); };
const mapLabelPoint = (geometry: any, bounds: { minX: number; minY: number; scale: number } | null) => { const points = positionsOf(geometry?.coordinates); if (!points.length || !bounds) return null; const [lng, lat] = points.reduce(([sumX, sumY], point) => [sumX + point[0], sumY + point[1]], [0, 0]).map((value) => value / points.length); return { x: 24 + (lng - bounds.minX) * bounds.scale, y: 356 - (lat - bounds.minY) * bounds.scale }; };

function FilterField({ field, options, selected, onChange, formatOption, fixedSelection = false }: { field: FilterConfig; options: string[]; selected: string; onChange: (value: string) => void; formatOption?: (value: string) => string; fixedSelection?: boolean }) {
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);
  const label = formatOption ?? ((value) => value);
  const visible = options.filter((option) => normalized(label(option)).includes(normalized(query)));
  const select = (value: string) => { onChange(value); setQuery(""); setOpen(false); };
  return <div className="filter-field"><label>{field.label}</label><Popover open={open} onOpenChange={(nextOpen) => { setOpen(nextOpen); if (!nextOpen) setQuery(""); }}><PopoverTrigger asChild><button type="button" className="filter-trigger" aria-label={fixedSelection ? `Desplegar filtro ${field.label}. Putumayo permanece seleccionado en la consulta territorial` : `Desplegar filtro ${field.label}`}><span>{selected === ALL ? "Seleccionar" : label(selected)}</span><ChevronDown size={16} /></button></PopoverTrigger><PopoverContent align="start" side="bottom" sideOffset={-1} className="filter-popover">{!fixedSelection && <div className="filter-search"><Search size={15} /><input autoFocus={open} aria-label={`Buscar ${field.label}`} value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Buscar" /></div>}{selected !== ALL && !fixedSelection && <button className="filter-option clear-option" onClick={() => select(ALL)}>Quitar selección</button>}<div className="filter-options">{visible.map((option) => <button key={option} className={selected === option ? "filter-option selected" : "filter-option"} onClick={() => select(option)}>{label(option)}</button>)}{!visible.length && <span className="no-options">Sin coincidencias</span>}</div></PopoverContent></Popover></div>;
}

function Heading({ number, title, detail, icon: Icon }: { number: string; title: string; detail?: string; icon: typeof Filter }) {
  return <div className="section-heading"><span className="section-number">{number}</span><div><div className="section-title"><Icon size={18} /><h2>{title}</h2></div>{detail && <p>{detail}</p>}</div></div>;
}

function MunicipalMap({ features, values, selected, years, activeYear, onSelect, onYearChange }: { features: MunicipalFeature[]; values: Map<string, { value: number | null; unit: string }>; selected: string; years: number[]; activeYear: string; onSelect: (name: string) => void; onYearChange: (year: string) => void }) {
  const bounds = useMemo(() => { const points = features.flatMap((item) => positionsOf(item.geometry.coordinates)); if (!points.length) return null; const xs = points.map((point) => point[0]); const ys = points.map((point) => point[1]); return { minX: Math.min(...xs), minY: Math.min(...ys), scale: Math.min(570 / (Math.max(...xs) - Math.min(...xs) || 1), 330 / (Math.max(...ys) - Math.min(...ys) || 1)) }; }, [features]);
  const [hovered, setHovered] = useState<string | null>(null);
  const numbers = Array.from(values.values()).map((entry) => entry.value).filter((value): value is number => value !== null);
  const low = Math.min(...numbers, 0); const high = Math.max(...numbers, 1);
  const tone = (value: number | null) => value === null ? "#e7eeee" : `hsl(${158 - (value - low) / (high - low || 1) * 132} 48% ${76 - (value - low) / (high - low || 1) * 31}%)`;
  return <article className="card municipal-map-card"><div className="card-heading"><div><p className="card-kicker">Resultado territorial</p><h3>Municipios del Putumayo</h3></div><div className="map-actions"><span className="map-legend">Seleccione un municipio</span><button className={selected !== ALL ? "map-reset" : "map-reset map-reset-placeholder"} onClick={() => onSelect(ALL)} disabled={selected === ALL}>Limpiar</button></div></div><div className="map-year-tabs" aria-label="Seleccionar vigencia">{years.map((year) => <button key={year} className={String(year) === activeYear ? "active" : ""} onClick={() => onYearChange(String(year))}>{year}</button>)}</div><svg className="municipal-map" viewBox="0 0 620 390" role="img" aria-label="Mapa interactivo de municipios del Putumayo">{features.map((item) => { const name = text(item.properties.name); const entry = values.get(normalized(name)); const value = entry?.value ?? null; const unit = entry?.unit || "Resultado"; const point = mapLabelPoint(item.geometry, bounds); const callout = mapCallouts[normalized(name)]; const isHovered = hovered === name; const displayUnit = normalized(unit) === "porcentaje" ? "%" : unit; const label = isHovered ? `${name}: ${value === null ? "Sin dato" : fmt(value)} ${displayUnit}` : name; return <g key={name} className="municipality-group" onMouseEnter={() => setHovered(name)} onMouseLeave={() => setHovered(null)} onClick={() => onSelect(name)}><path d={mapPath(item.geometry, bounds)} fill={tone(value)} className={normalized(selected) === normalized(name) ? "municipality active" : "municipality"} />{point && callout && <line className="municipality-callout" x1={point.x} y1={point.y} x2={point.x + callout.dx} y2={point.y + callout.dy} />}{point && <text x={point.x + (callout?.dx || 0)} y={point.y + (callout?.dy || 0)} className={isHovered ? "municipality-label active" : "municipality-label"} textAnchor={callout ? callout.dx < 0 ? "end" : callout.dx > 0 ? "start" : "middle" : "middle"}>{label}</text>}<title>{`${name}: ${value === null ? "sin dato" : `${fmt(value)} ${displayUnit}`}`}</title></g>; })}</svg><div className="map-scale"><span>Sin dato</span><i /><span>Mayor resultado</span></div></article>;
}

function OdsPanel({ objective, goal }: { objective: string; goal: string }) { return <article className="card ods-card ods-card-inline"><p className="card-kicker">Agenda 2030</p><h3><strong>ODS:</strong> {objective || "No disponible"}</h3>{hasMeaningfulValue(goal) && <p><strong>Meta ODS:</strong> {goal}</p>}</article>; }
function VariationPanel({ variations }: { variations: Variation[] }) { return <article className="card variations-card"><div className="card-heading"><div><p className="card-kicker">Variación anual</p><h3>Trayectoria por serie</h3></div><span className="formula-note">Diferencia y % relativo</span></div><div className="variations-scroll">{variations.map((item) => <div className="variation-series" key={item.name}><strong>{item.name}</strong><div>{item.items.map((entry) => <span key={entry.label}><b>{entry.label}</b><em className={entry.difference === null ? "neutral" : entry.difference >= 0 ? "positive" : "negative"}>{entry.difference === null ? "N/D" : `${entry.difference >= 0 ? "+" : ""}${fmt(entry.difference)}`}</em><small>{entry.relative === null ? "Relativa: no calculable" : `Relativa: ${entry.relative >= 0 ? "+" : ""}${fmt(entry.relative)}%`}</small></span>)}</div></div>)}</div></article>; }

function EvolutionPanel({ years, unit, chartData, activeSeries, history, onHistoryChange, showAll, onShowAll, filterContext }: { years: number[]; unit: string; chartData: Record<string, string | number | null>[]; activeSeries: Series[]; history: boolean; onHistoryChange: (value: boolean) => void; showAll: boolean; onShowAll: () => void; filterContext: string }) {
  return <article className="chart-card"><div className="chart-controls"><fieldset><legend>Vista de serie</legend><label><input type="radio" checked={history} onChange={() => onHistoryChange(true)} />Serie histórica completa</label><label><input type="radio" checked={!history} onChange={() => onHistoryChange(false)} />Solo vigencia seleccionada</label></fieldset>{activeSeries.length > 6 && <button className="button button-quiet" onClick={onShowAll}>{showAll ? "Reducir series" : `Ver las ${activeSeries.length} series`}</button>}</div><div className="chart-label"><span>Eje Y</span><strong>Resultado ({unit})</strong><em>Barras por vigencia y línea de tendencia</em></div><div className="chart-wrapper">{chartData.length && years.length ? <ResponsiveContainer width="100%" height={390}><ComposedChart data={chartData} margin={{ top: 28, right: 28, left: 12, bottom: 6 }}><CartesianGrid vertical={false} stroke="#d8e1e3" strokeDasharray="2 4" /><XAxis dataKey="year" tickLine={false} axisLine={{ stroke: "#9caeb4" }} /><YAxis tickLine={false} axisLine={false} width={62} /><Tooltip contentStyle={{ borderRadius: 0, border: "1px solid #d1dddf", boxShadow: "0 12px 28px rgba(0,72,132,.12)" }} formatter={(value: number) => [`${fmt(value)}${normalized(unit) === "porcentaje" ? " %" : ""}`, "Resultado"]} />{activeSeries.map((item, index) => <Bar key={`${item.name}-bar`} dataKey={item.name} fill={colors[index % colors.length]} fillOpacity={.22} stroke={colors[index % colors.length]} strokeOpacity={.52} barSize={activeSeries.length === 1 ? 30 : 14} radius={[2, 2, 0, 0]} />)}{activeSeries.map((item, index) => <Line key={`${item.name}-line`} type="monotone" dataKey={item.name} stroke={colors[index % colors.length]} strokeWidth={normalized(item.name).startsWith("putumayo") ? 3 : 2} dot={{ r: 3, strokeWidth: 1 }} activeDot={{ r: 5 }} connectNulls><LabelList dataKey={item.name} position="top" offset={8} fill={colors[index % colors.length]} fontSize={9} fontWeight={700} formatter={(value: Value) => value === null || value === undefined ? "" : `${fmt(value)}${normalized(unit) === "porcentaje" ? "%" : ""}`} /></Line>)}</ComposedChart></ResponsiveContainer> : <div className="empty-inline">No hay datos suficientes para construir la serie histórica.</div>}</div><p className="chart-caption">Las barras muestran el valor de cada vigencia y la línea continua permite reconocer la tendencia de la misma serie.</p><p className="chart-filter-context"><span>Filtros aplicados</span><strong>{filterContext}</strong></p></article>;
}

export default function Home() {
  const rows = EMBEDDED_DATA.rows;
  const technicalProfiles = useMemo(() => new Map(EMBEDDED_DATA.technicalProfiles.map((profile) => [profile.code, profile])), []);
  const workbookStructure = EMBEDDED_DATA.workbookStructure;
  const [filters, setFilters] = useState<Filters>({});
  const [contrast, setContrast] = useState(false);
  const [scale, setScale] = useState(1);
  const [metadata, setMetadata] = useState(false);
  const [history, setHistory] = useState(true);
  const [showAll, setShowAll] = useState(false);
  const [techOpen, setTechOpen] = useState(true);
  const [lastFilterKey, setLastFilterKey] = useState("");
  const [restrictIndicatorOptions, setRestrictIndicatorOptions] = useState(false);
  const [query, setQuery] = useState("");
  const [page, setPage] = useState(1);
  const [activeNav, setActiveNav] = useState("filtros");
  const [municipalShapes, setMunicipalShapes] = useState<MunicipalFeature[]>([]);
  const specialIndex = useMemo(() => createSpecialIndex(rows), [rows]);

  useEffect(() => { fetch(MUNICIPAL_MAP_URL).then((response) => response.json()).then((topology) => { const collection = topoFeature(topology, topology.objects.mpios) as unknown as { features: MunicipalFeature[] }; setMunicipalShapes(collection.features.filter((item) => normalized(item.properties.dpt) === "putumayo")); }).catch(() => setMunicipalShapes([])); }, []);
  useEffect(() => { const escape = (event: KeyboardEvent) => event.key === "Escape" && setMetadata(false); window.addEventListener("keydown", escape); return () => window.removeEventListener("keydown", escape); }, []);
  useEffect(() => { const targets = ["filtros", "resultados", "evolucion", "ficha", "tabla"].map((id) => ({ id, node: document.getElementById(id) })).filter((item): item is { id: string; node: HTMLElement } => Boolean(item.node)); const observer = new IntersectionObserver((entries) => { const visible = entries.filter((entry) => entry.isIntersecting).sort((a, b) => b.intersectionRatio - a.intersectionRatio)[0]; if (visible) { const target = targets.find((item) => item.node === visible.target); if (target) setActiveNav(target.id); } }, { rootMargin: "-18% 0px -64% 0px", threshold: [0, .25, .6] }); targets.forEach((item) => observer.observe(item.node)); return () => observer.disconnect(); }, [rows, filters]);

  const change = (key: string, value: string) => {
    if (key === DEPARTMENT_KEY && value === ALL) return;
    if (key === "INDICADOR" && value !== ALL) setRestrictIndicatorOptions(false);
    if (key === "AREA TEMATICA") setRestrictIndicatorOptions(value !== ALL);
    if (key === "TEMA" && value !== ALL) setRestrictIndicatorOptions(true);
    if (["AREA TEMATICA", "TEMA", "INDICADOR"].includes(key) && value === ALL) setRestrictIndicatorOptions(false);
    setFilters((current) => {
      if (["AREA TEMATICA", "TEMA", "INDICADOR"].includes(key) && value === ALL) {
        return emptyFilters();
      }
      if (key === "INDICADOR" && value !== ALL) {
        const codeRows = rows.filter((row) => rowValue(row, "ID_INDICADOR") === value && numeric(row.RESULTADO) !== null);
        const resetDimensions = Object.fromEntries(dimensions.map((field) => [field.key, ALL]));
        const representative = codeRows[0];
        const latestYear = distinct(codeRows.map((row) => rowValue(row, "VIGENCIA"))).map((year) => numeric(year)).filter((year): year is number => year !== null).sort((a, b) => b - a)[0];
        return { ...current, ...resetDimensions, [DEPARTMENT_KEY]: "Putumayo", "AREA TEMATICA": valueFor(representative, "AREA TEMATICA", specialIndex) || ALL, TEMA: valueFor(representative, "TEMA", specialIndex) || ALL, INDICADOR: value, VIGENCIA: latestYear === undefined ? ALL : String(latestYear) };
      }
      if (key === "AREA TEMATICA") return { ...current, [key]: value, TEMA: ALL, INDICADOR: ALL, VIGENCIA: ALL, [DEPARTMENT_KEY]: ALL, ...Object.fromEntries(dimensions.map((field) => [field.key, ALL])) };
      if (key === DEPARTMENT_KEY) return { ...current, [key]: value, [MUNICIPALITY_KEY]: ALL };
      if (key === "TEMA") {
        const topicRows = value === ALL ? [] : rows.filter((row) => normalized(valueFor(row, "TEMA", specialIndex)) === normalized(value));
        const representative = topicRows[0];
        return { ...current, "AREA TEMATICA": value === ALL ? current["AREA TEMATICA"] : valueFor(representative, "AREA TEMATICA", specialIndex) || ALL, TEMA: value, INDICADOR: ALL, VIGENCIA: ALL, [DEPARTMENT_KEY]: ALL, ...Object.fromEntries(dimensions.map((field) => [field.key, ALL])) };
      }
      return { ...current, [key]: value };
    });
    if (key === "INDICADOR" || ["AREA TEMATICA", "TEMA", "VIGENCIA"].includes(key)) {
      if (key !== "VIGENCIA") setLastFilterKey("");
    } else if (variationDimensionKeys.has(key) && value !== ALL) {
      setLastFilterKey(key);
    } else if (variationDimensionKeys.has(key)) {
      setLastFilterKey("");
    }
    setPage(1);
  };

  const selectedCode = filters.INDICADOR ?? ALL;
  const technicalProfile = technicalProfiles.get(selectedCode);
  const indicatorLabels = useMemo(() => new Map(distinct(rows.map((row) => rowValue(row, "ID_INDICADOR"))).map((code) => { const names = distinct(rows.filter((row) => rowValue(row, "ID_INDICADOR") === code).map((row) => rowValue(row, "INDICADOR"))); const profileName = technicalProfiles.get(code)?.fields["Nombre del Indicador"]; const base = profileName || sharedStart(names).replace(/[\s\-–—:;(]+$/, "").trim() || names[0] || "Indicador"; return [code, `${code.replace(/^IND-DANE-/i, "")} - ${base}`]; })), [rows, technicalProfiles]);
  const optionLabel = (key: string, value: string) => key === "INDICADOR" ? indicatorLabels.get(value) ?? value : value;
  const options = (key: string) => { const filterScope = key === DEPARTMENT_KEY ? { ...filters, [DEPARTMENT_KEY]: ALL, [MUNICIPALITY_KEY]: ALL } : key === "INDICADOR" && !restrictIndicatorOptions ? emptyFilters() : filters; const scoped = matchesFilters(rows, filterScope, specialIndex, key).filter((row) => key !== "VIGENCIA" || numeric(row.RESULTADO) !== null); const values = distinct(scoped.map((row) => valueFor(row, key, specialIndex))); const withoutPrincipal = key === SPECIAL_KEY && PRINCIPAL_DIRECT_CODES.has(selectedCode) ? values.filter((value) => normalized(value) !== "total") : values; return withoutPrincipal.sort((a, b) => key === "VIGENCIA" ? Number(a) - Number(b) : optionLabel(key, a).localeCompare(optionLabel(key, b), "es")); };
  const hasSelection = Object.values(filters).some((value) => value && value !== ALL);
  const hasIndicator = selectedCode !== ALL;
  const hasSelectedYear = filters.VIGENCIA !== ALL;
  const activeQuery = hasIndicator && hasSelectedYear;
  const awaitingSpecialVariant = activeQuery && VARIANT_REQUIRED_CODES.has(selectedCode) && Boolean(specialIndex.get(selectedCode)?.size) && filters[SPECIAL_KEY] === ALL;
  const awaitingDirectDimension = activeQuery && selectedCode === DIRECT_DIMENSION_CODE && filters["NIVEL_EDUCATIVO_GENERAL"] === ALL;
  const canDisplayResults = activeQuery && !awaitingSpecialVariant && !awaitingDirectDimension;
  const onlyPrincipalRows = (items: Row[]) => PRINCIPAL_DIRECT_CODES.has(selectedCode) && filters[SPECIAL_KEY] === ALL ? items.filter((row) => normalized(valueFor(row, SPECIAL_KEY, specialIndex)) === "total") : items;
  const filtered = useMemo(() => onlyPrincipalRows(matchesFilters(rows, filters, specialIndex)), [rows, filters, specialIndex, selectedCode]);
  const withoutYear = useMemo(() => onlyPrincipalRows(matchesFilters(rows, { ...filters, VIGENCIA: ALL }, specialIndex)), [rows, filters, specialIndex, selectedCode]);
  const indicator = filtered[0] ?? withoutYear[0];
  const code = rowValue(indicator, "ID_INDICADOR");
  const unit = rowValue(indicator, "UNIDAD_MEDIDA") || "Resultado";
  const years = useMemo(() => distinct(withoutYear.filter((row) => numeric(row.RESULTADO) !== null).map((row) => rowValue(row, "VIGENCIA"))).map(Number).filter(Number.isFinite).sort((a, b) => a - b), [withoutYear]);
  const visibleDimensions = activeQuery ? dimensions.filter((field) => {
    if (field.key === SPECIAL_KEY && !specialIndex.get(selectedCode)?.size) return false;
    const marked = technicalProfile?.markedDimensions ?? [];
    const permittedByFicha = field.key === DEPARTMENT_KEY || (selectedCode === "IND-DANE-064" && field.key === "NIVEL_EDUCATIVO SUPERIOR") || (selectedCode === DIRECT_DIMENSION_CODE ? field.key === "NIVEL_EDUCATIVO_GENERAL" : !marked.length || marked.includes(field.key));
    const available = options(field.key);
    return permittedByFicha && (field.key === DEPARTMENT_KEY || available.length > 0) && (field.key === SPECIAL_KEY || field.key === DEPARTMENT_KEY || field.key === MUNICIPALITY_KEY || (field.key === "NIVEL_EDUCATIVO SUPERIOR" && selectedCode === "IND-DANE-064") || available.length > 1);
  }) : [];
  const lineLabel = (row: Row) => { const territory = valueFor(row, MUNICIPALITY_KEY, specialIndex) || valueFor(row, DEPARTMENT_KEY, specialIndex) || "Resultado"; const special = valueFor(row, SPECIAL_KEY, specialIndex); return [territory, filters[SPECIAL_KEY] === ALL && special ? special : ""].filter(Boolean).join(" · "); };
  const dataForChart = history ? withoutYear : filtered;
  const hasAdditionalDisaggregation = dimensions.some((field) => ![DEPARTMENT_KEY, MUNICIPALITY_KEY, SPECIAL_KEY].includes(field.key) && filters[field.key] !== ALL) || filters[SPECIAL_KEY] !== ALL;
  const preferDirectTotal = normalized(filters[DEPARTMENT_KEY]) === "putumayo" && (DIRECT_TERRITORIAL_CODES.has(selectedCode) || !hasAdditionalDisaggregation);
  const series = useMemo<Series[]>(() => aggregateSeries(dataForChart, lineLabel, specialIndex, preferDirectTotal).sort((a, b) => a.name.localeCompare(b.name, "es")), [dataForChart, filters, specialIndex, preferDirectTotal]);
  const activeSeries = useMemo(() => { if (showAll || series.length <= 6) return series; const territorial = series.find((item) => normalized(item.name).startsWith("putumayo")); return territorial ? [territorial, ...series.filter((item) => item !== territorial).slice(0, 5)] : series.slice(0, 6); }, [series, showAll]);
  const chartData = years.map((year) => Object.fromEntries([["year", year], ...activeSeries.map((item) => [item.name, item.points.get(year) ?? null])])) as Record<string, string | number | null>[];
  const mainRows = useMemo(() => { const department = filters[DEPARTMENT_KEY] === ALL ? "putumayo" : filters[DEPARTMENT_KEY]; const municipality = filters[MUNICIPALITY_KEY]; const departmentRows = filtered.filter((row) => normalized(valueFor(row, DEPARTMENT_KEY, specialIndex)) === normalized(department)); const municipalityRows = municipality !== ALL ? filtered.filter((row) => normalized(valueFor(row, MUNICIPALITY_KEY, specialIndex)) === normalized(municipality)) : departmentRows; if (municipality !== ALL) return municipalityRows; if (departmentRows.length || DIRECT_TERRITORIAL_CODES.has(selectedCode)) return departmentRows; return filtered; }, [filtered, filters, specialIndex, selectedCode]);
  const mainResult = useMemo(() => aggregateResult(mainRows, specialIndex, preferDirectTotal), [mainRows, specialIndex, preferDirectTotal]);
  const main = mainResult.row;
  const variationLabel = (row: Row) => lastFilterKey && filters[lastFilterKey] !== ALL ? valueFor(row, lastFilterKey, specialIndex) || "Putumayo" : "Putumayo";
  const variationRows = useMemo(() => {
    if (lastFilterKey && filters[lastFilterKey] !== ALL) return withoutYear;
    const putumayoRows = withoutYear.filter((row) => normalized(valueFor(row, DEPARTMENT_KEY, specialIndex)) === "putumayo" && !valueFor(row, MUNICIPALITY_KEY, specialIndex));
    const hasTotal = putumayoRows.some((row) => normalized(valueFor(row, SPECIAL_KEY, specialIndex)) === "total");
    return hasTotal ? putumayoRows.filter((row) => normalized(valueFor(row, SPECIAL_KEY, specialIndex)) === "total") : putumayoRows;
  }, [withoutYear, filters, lastFilterKey, specialIndex]);
  const variationSeries = useMemo<Series[]>(() => aggregateSeries(variationRows, variationLabel, specialIndex, preferDirectTotal), [variationRows, filters, lastFilterKey, specialIndex, preferDirectTotal]);
  const variations: Variation[] = variationSeries.map((item) => ({ name: item.name, items: years.slice(1).map((year) => { const before = item.points.get(year - 1); const now = item.points.get(year); const difference = before === undefined || now === undefined ? null : now - before; const relative = difference === null || !before ? null : difference / before * 100; return { label: `${year - 1}–${year}`, difference, relative }; }) }));
  const mapRows = useMemo(() => selectedCode === ALL || SPECIAL_DIRECT_CODES.has(selectedCode) || (specialIndex.get(selectedCode)?.size ?? 0) > 1 && filters[SPECIAL_KEY] === ALL ? [] : matchesFilters(rows, { ...filters, [MUNICIPALITY_KEY]: ALL }, specialIndex), [rows, filters, selectedCode, specialIndex]);
  const municipalityValues = useMemo(() => {
    const grouped = new Map<string, Row[]>();
    mapRows.forEach((row) => { const territory = normalized(valueFor(row, MUNICIPALITY_KEY, specialIndex)); if (territory) grouped.set(territory, [...(grouped.get(territory) ?? []), row]); });
    return new Map(Array.from(grouped.entries()).map(([territory, territoryRows]) => { const result = aggregateResult(territoryRows, specialIndex, false); return [territory, { value: result.value, unit: rowValue(result.row, "UNIDAD_MEDIDA") || "Resultado" }]; }));
  }, [mapRows, specialIndex]);
  const hasMunicipalData = municipalShapes.some((shape) => municipalityValues.has(normalized(shape.properties.name)));
  const odsObjective = valueFor(indicator, "OBJETIVO ODS", specialIndex);
  const odsGoal = valueFor(indicator, "META ODS", specialIndex);
  const hasOds = [odsObjective, odsGoal].some((value) => value && !["no disponible", "n/a", "na", "-"].includes(normalized(value)));
  const dataColumns = useMemo(() => dataOrder.filter((key) => rows.some((row) => rowValue(row, key))), [rows]);
  const metadataSummary = useMemo(() => {
    const values = (key: string) => distinct(rows.map((row) => rowValue(row, key)));
    const periods = values("VIGENCIA").map((value) => numeric(value)).filter((value): value is number => value !== null).sort((a, b) => a - b);
    return {
      records: rows.length, indicators: values("ID_INDICADOR").length, fields: dataColumns.length, periods: periods.length ? `${periods[0]}–${periods[periods.length - 1]}` : "No disponible",
      themes: values("TEMA").length, subthemes: values("SUBTEMA").length, departments: values(DEPARTMENT_KEY).length, territories: 14, municipalities: 13, divipola: values("CODIDO DIVIPOLA").length,
      units: values("UNIDAD_MEDIDA").length, sources: values("FUENTE").length, dates: values("FECHA ULTIMA CONSULTA"),
      odsRecords: rows.filter((row) => Boolean(valueFor(row, "OBJETIVO ODS", specialIndex) || valueFor(row, "META ODS", specialIndex))).length,
      numericResults: rows.filter((row) => numeric(row.RESULTADO) !== null).length, sourcesWithLink: rows.filter((row) => isWebLink(rowValue(row, "LINK FUENTE"))).length,
    };
  }, [rows, dataColumns, specialIndex]);
  const metadataStructure = useMemo(() => {
    const main = workbookStructure.find((sheet) => normalized(sheet.name) === normalized(MAIN_SHEET));
    return { sheets: workbookStructure.length, technicalSheets: workbookStructure.filter((sheet) => /^id\s*\d+/i.test(text(sheet.name))).length, mainSheet: main?.name || MAIN_SHEET, mainRecords: main?.rowCount || rows.length };
  }, [workbookStructure, rows.length]);
  const resultFilterContext = useMemo(() => distinct([DEPARTMENT_KEY, MUNICIPALITY_KEY, "ÁREA_GEOGRÁFICA", "SEXO", "NIVEL_EDUCATIVO_GENERAL", "NIVEL_EDUCATIVO SUPERIOR", SPECIAL_KEY].map((key) => filters[key]).filter((value) => Boolean(value) && value !== ALL)).join(" · ") || "Putumayo", [filters]);
  const metadataGroups = useMemo(() => metadataFieldGroups.map((group) => ({ ...group, fields: group.fields.filter((field) => dataColumns.includes(field)) })).filter((group) => group.fields.length), [dataColumns]);
  const tableRows = rows.filter((row) => !query || dataColumns.some((key) => normalized(rowValue(row, key)).includes(normalized(query))));
  const perPage = 10; const totalPages = Math.max(1, Math.ceil(tableRows.length / perPage)); const currentRows = tableRows.slice((page - 1) * perPage, page * perPage);
  const clearPrimaryFilters = () => { setFilters(emptyFilters()); setLastFilterKey(""); setRestrictIndicatorOptions(false); setPage(1); };
  const clearDimensions = () => {
    setFilters((current) => clearDimensionSelections(current, dimensions.map((field) => field.key), DEPARTMENT_KEY, ALL));
    setLastFilterKey("");
    setPage(1);
  };
  const download = () => { const dataSheet = XLSX.utils.json_to_sheet(rows.map((row) => Object.fromEntries(dataColumns.map((key) => [labels[key] ?? key, rowValue(row, key)])))); const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, dataSheet, "Datos"); XLSX.writeFile(book, "bateria_indicadores_dane_putumayo_completo.xlsx"); };
  const technical = activeQuery && indicator ? (() => {
    const profile = technicalProfile?.fields ?? {};
    const profileName = profile["Nombre del Indicador"];
    const profileDescription = profile["Descripcion del Inidcador"] || profile["Descripción del Indicador"] || profile["Descripcion del Indicador"];
    const municipalities = distinct(withoutYear.map((row) => valueFor(row, MUNICIPALITY_KEY, specialIndex)));
    const reportedDimensions = visibleDimensions.map((field) => field.label);
    const base: Record<string, string> = {
      "Código": code,
      "Nombre": profileName || rowValue(indicator, "INDICADOR"),
      "Descripción": profileDescription || rowValue(indicator, "DESCRIPCION") || "No disponible",
      "Unidad": profile["Unidad de Medida"] || unit,
      "Periodicidad": profile["Periocidad de Medición"] || profile["Periodicidad de Medición"] || rowValue(indicator, "PERIODICIDAD_REPORTE") || "No disponible",
      "Fuente": profile["Fuente de la Información"] || rowValue(indicator, "FUENTE") || "No disponible",
      "Enlace de fuente": rowValue(indicator, "LINK FUENTE"),
      "Periodo de referencia": years.length ? `${years[0]}-${years[years.length - 1]}` : profile["Periodo de referencia"] || "No disponible",
      "Cobertura geográfica": profile["Cobertura geografica"] || distinct(withoutYear.map((row) => valueFor(row, DEPARTMENT_KEY, specialIndex))).join(" · ") || "No disponible",
      "Fecha de última consulta": formatDate(rowValue(indicator, "FECHA ULTIMA CONSULTA")),
    };
    if (municipalities.length) base["Municipios disponibles"] = municipalities.join(" · ");
    if (reportedDimensions.length) base["Desagregaciones reportadas"] = reportedDimensions.join(" · ");
    const objective = valueFor(indicator, "OBJETIVO ODS", specialIndex);
    if (hasMeaningfulValue(objective)) base["ODS"] = objective;
    ["Formula de Calculo", "Nombre de la publicación u operación estadística", "Nomenclaturas y clasificaciones estadísticas", "Alcance temático", "Población objetivo", "Línea Base del Indicador", "Año Línea Base del Indicador"].forEach((name) => { if (hasMeaningfulValue(profile[name])) base[name] = profile[name]; });
    const meta = valueFor(indicator, "META ODS", specialIndex);
    if (hasMeaningfulValue(meta)) base["Meta ODS"] = meta;
    return base;
  })() : null;


  const appStyle = {
    fontSize: `${scale}rem`,
    "--asset-symbol": `url("${SYMBOL_URL}")`,
    "--asset-texture": `url("${TEXTURE_URL}")`,
  } as CSSProperties;

  return (
    <div className={contrast ? "app high-contrast" : "app"} style={appStyle}>
      <a className="skip-link" href="#contenido">Saltar al contenido principal</a>
      <header className="top-strip"><div className="top-strip-content"><span>Gobernación del Putumayo</span><span>Información estadística para decisiones públicas</span></div></header>
      <div className="hero" style={{ backgroundImage: `linear-gradient(90deg,rgba(246,249,250,.98) 0%,rgba(246,249,250,.95) 48%,rgba(246,249,250,.28) 100%),url(${LANDSCAPE_URL})`, height: "150px" }}>
        <div className="hero-inner"><div className="identity-block"><div className="official-logo"><img src={LOGO_URL} alt="Logo de la Gobernación del Putumayo" /></div><div className="identity-copy"><p className="eyebrow">Secretaría de Planeación · Área de Estadística</p><h1>Batería <em>indicadores DANE</em></h1><p className="hero-description">Consulta, compara y descarga los resultados disponibles para el departamento del Putumayo.</p></div></div><div className="hero-brand-mark"><img src={SYMBOL_URL} alt="" /><div><span>Instrumento territorial</span><strong>Batería DANE<br />Putumayo</strong></div></div></div>
      </div>
      <div className="app-layout" id="contenido">
        <aside className="context-rail" aria-label="Navegación de secciones"><div className="rail-sticky"><div className="rail-brand-lockup"><img src={SYMBOL_URL} alt="" className="rail-symbol" /><div><span>Instrumento público</span><strong>Batería DANE</strong></div></div><p>Consulta territorial</p><nav><a className={activeNav === "filtros" ? "active" : ""} href="#filtros"><span>01</span>Filtros</a><a className={activeNav === "resultados" ? "active" : ""} href="#resultados"><span>02</span>Resultados</a><a className={activeNav === "evolucion" ? "active" : ""} href="#evolucion"><span>03</span>Evolución</a><a className={activeNav === "ficha" ? "active" : ""} href="#ficha"><span>04</span>Ficha técnica</a><a className={activeNav === "tabla" ? "active" : ""} href="#tabla"><span>05</span>Datos</a></nav><div className="rail-coordinates"><span>01°08′ N</span><span>76°39′ O</span><small>Plano territorial</small></div><div className="rail-source"><MapPin size={15} /><span>Putumayo, Colombia</span></div></div></aside>
        <main className="content-plane" style={{ backgroundImage: `linear-gradient(rgba(247,250,250,.92),rgba(247,250,250,.98)),url(${TEXTURE_URL})` }}>
          <section id="filtros" className="dashboard-section filter-section">
            <div className="filter-toolbar"><div className="filter-instrument-mark"><img src={SYMBOL_URL} alt="" /><div><span>Plano de consulta territorial</span><strong>Putumayo · datos para decisiones públicas</strong></div></div><div className="toolbar-actions"><button className="button button-primary button-metadata" onClick={() => setMetadata(true)}><Info size={18} />Ver metadatos</button><button className="button button-quiet button-clean-primary" onClick={clearPrimaryFilters}>Limpiar filtros principales</button></div></div>
            <div className="filter-grid">{primary.map((field) => <FilterField key={field.key} field={field} options={options(field.key)} selected={filters[field.key] ?? ALL} onChange={(value) => change(field.key, value)} formatOption={(value) => optionLabel(field.key, value)} />)}</div>
            {visibleDimensions.length > 0 && <div className="inline-dimensions"><div className="dimensions-toolbar"><Heading number="01B" title="Desagregaciones" icon={SlidersHorizontal} /><button className="button button-quiet button-clean-dimensions" onClick={clearDimensions}>Limpiar desagregaciones</button></div><div className="filter-grid dimensions-grid">{visibleDimensions.map((field) => <FilterField key={field.key} field={field} options={options(field.key)} selected={filters[field.key] ?? ALL} onChange={(value) => change(field.key, value)} fixedSelection={field.key === DEPARTMENT_KEY} />)}</div></div>}
          </section>
          <section id="resultados" className="dashboard-section">
            <Heading number="02" title="Resultado del indicador" detail="Consulte los resultados oficiales según territorio, vigencia y desagregación disponibles." icon={BarChart3} />
            {!hasIndicator ? <div className="empty-state"><CircleHelp size={28} /><h3>Seleccione un indicador para iniciar la consulta</h3><p>Puede elegirlo directamente o acotar la búsqueda por área temática y tema.</p></div> : !hasSelectedYear ? <div className="empty-state"><CircleHelp size={28} /><h3>Seleccione una vigencia para consultar el resultado</h3><p>Elija un año en el filtro Vigencia. Las desagregaciones, resultados y ficha técnica se activarán con esta selección.</p></div> : awaitingSpecialVariant ? <div className="empty-state"><CircleHelp size={28} /><h3>Seleccione una desagregación especial</h3><p>Este indicador se reporta por variantes. Elija una opción en Desagregaciones para consultar su resultado.</p></div> : awaitingDirectDimension ? <div className="empty-state"><CircleHelp size={28} /><h3>Seleccione el nivel educativo</h3><p>Este indicador se reporta por nivel educativo general. Elija una opción en Desagregaciones para consultar su resultado.</p></div> : !filtered.length ? <div className="empty-state"><CircleHelp size={28} /><h3>No hay registros para los filtros seleccionados</h3><p>Elimine la última selección o use Limpiar filtros principales para iniciar una nueva consulta.</p><button className="button button-primary" onClick={clearPrimaryFilters}>Limpiar filtros principales</button></div> : <>
              <div className={hasOds ? "result-grid" : "result-grid no-ods"}>
                <article className="card indicator-card"><p className="card-kicker">Identificación</p><span className="indicator-code">{code}</span><h3>{technical?.Nombre || rowValue(indicator, "INDICADOR")}</h3><div className="card-footer-text">{valueFor(indicator, "AREA TEMATICA", specialIndex)} <span>·</span> {valueFor(indicator, "TEMA", specialIndex)}</div></article>
                <article className="card main-result-card"><div className="result-header"><p className="card-kicker">Resultado principal</p><span>{resultFilterContext}</span></div><div className="big-number">{fmt(mainResult.value)}</div><p className="result-unit">{normalized(unit) === "porcentaje" ? "%" : unit}</p><p className="result-context">Vigencia <strong>{filters.VIGENCIA === ALL ? "seleccionada" : filters.VIGENCIA}</strong></p><div className="result-meta-list"><span><b>DIVIPOLA</b>{valueFor(main, "CODIGO DIVIPOLA", specialIndex) || "No disponible"}</span>{rowValue(main, "ESTADO") && <span className="status-pill"><b>Estado</b>{rowValue(main, "ESTADO")}</span>}</div><div className="source-link"><span>Fuente</span>{isWebLink(rowValue(indicator, "LINK FUENTE")) ? <a href={rowValue(indicator, "LINK FUENTE")} target="_blank" rel="noreferrer">{short(rowValue(indicator, "FUENTE"), 58)}</a> : <strong>{short(rowValue(indicator, "FUENTE"), 58)}</strong>}</div></article>
                {hasOds && <OdsPanel objective={odsObjective} goal={odsGoal} />}
              </div>
              <div className="result-bottom-row">
                <div className="territorial-slot">{hasMunicipalData ? <MunicipalMap features={municipalShapes} values={municipalityValues} selected={filters[MUNICIPALITY_KEY]} years={years} activeYear={filters.VIGENCIA} onSelect={(name) => change(MUNICIPALITY_KEY, name)} onYearChange={(year) => change("VIGENCIA", year)} /> : <div id="evolucion" className="territorial-evolution"><div className="card-heading"><div><p className="card-kicker">Resultado histórico</p><h3>Evolución histórica del indicador</h3></div></div><EvolutionPanel years={years} unit={unit} chartData={chartData} activeSeries={activeSeries} history={history} onHistoryChange={setHistory} showAll={showAll} onShowAll={() => setShowAll(!showAll)} filterContext={resultFilterContext} /></div>}</div>
                <div className="variation-slot"><VariationPanel variations={variations} /></div>
              </div>
            </>}
          </section>
          {canDisplayResults && hasMunicipalData && <section id="evolucion" className="dashboard-section chart-section"><Heading number="03" title="Evolución histórica del indicador" detail="Cada punto muestra el resultado de su vigencia; no se requiere pasar el cursor para consultarlo." icon={LineChartIcon} /><EvolutionPanel years={years} unit={unit} chartData={chartData} activeSeries={activeSeries} history={history} onHistoryChange={setHistory} showAll={showAll} onShowAll={() => setShowAll(!showAll)} filterContext={resultFilterContext} /></section>}
          <section id="ficha" className="dashboard-section technical-section"><Heading number="04" title="Ficha técnica" detail="Ficha técnica institucional del indicador seleccionado." icon={BookOpen} />{technical ? <article className="technical-card"><button className="technical-toggle" onClick={() => setTechOpen(!techOpen)} aria-expanded={techOpen}><div><span className="technical-code">{code}</span><h3>{technical.Nombre}</h3></div>{techOpen ? <ChevronUp /> : <ChevronDown />}</button>{techOpen && <div className="technical-content"><div className="technical-intro"><p>{technical.Descripción}</p><div className="formula-box"><span>Fuente de información</span><strong>{technical.Fuente}</strong></div></div><div className="technical-grid">{Object.entries(technical).filter(([name]) => !["Código", "Nombre", "Descripción", "Fuente"].includes(name)).map(([name, value]) => <div key={name}><dt>{name}</dt><dd>{value.startsWith("http") ? <a href={value} target="_blank" rel="noreferrer">Abrir enlace</a> : value}</dd></div>)}</div></div>}</article> : <div className="empty-inline">La ficha técnica se mostrará al seleccionar un indicador.</div>}</section>
          <section id="tabla" className="dashboard-section table-section"><Heading number="05" title="Tabla completa de indicadores" detail="La tabla presenta el conjunto completo de registros disponibles, independiente de los filtros de consulta." icon={FileSpreadsheet} /><div className="table-controls"><strong className="table-total">{rows.length} registros disponibles</strong><div className="table-actions"><div className="global-search"><Search size={16} /><input value={query} onChange={(event) => { setQuery(event.target.value); setPage(1); }} placeholder="Buscar en la tabla completa" /></div><button className="button button-primary" onClick={download}><Download size={16} />Descargar datos completos</button></div></div><div className="data-table-wrap"><table className="data-table"><thead><tr>{dataColumns.map((key) => <th key={key}>{labels[key] ?? key}</th>)}</tr></thead><tbody>{currentRows.map((row, index) => <tr key={`${rowValue(row, "ID_INDICADOR")}-${index}`}>{dataColumns.map((key) => <td key={key}>{key === "LINK FUENTE" && isWebLink(rowValue(row, key)) ? <a href={rowValue(row, key)} target="_blank" rel="noreferrer">Consultar</a> : key === "RESULTADO" ? `${fmt(row[key])}${normalized(rowValue(row, "UNIDAD_MEDIDA")) === "porcentaje" ? " %" : ""}` : rowValue(row, key) || "—"}</td>)}</tr>)}</tbody></table></div><div className="pagination"><span>Mostrando {currentRows.length ? (page - 1) * perPage + 1 : 0}–{Math.min(page * perPage, tableRows.length)} de {tableRows.length} registros</span><div><button className="button button-quiet" disabled={page === 1} onClick={() => setPage(page - 1)}>Anterior</button><strong>Página {page} de {totalPages}</strong><button className="button button-quiet" disabled={page === totalPages} onClick={() => setPage(page + 1)}>Siguiente</button></div></div></section>
          <footer className="source-footer"><span>Fuente: <strong>Base de indicadores DANE</strong>.</span>{hasIndicator && isWebLink(rowValue(indicator, "LINK FUENTE")) && <a href={rowValue(indicator, "LINK FUENTE")} target="_blank" rel="noreferrer">Consultar fuente externa</a>}</footer>
        </main>
      </div>
      <aside className="accessibility-bar" aria-label="Herramientas de accesibilidad"><button aria-label="Cambiar contraste" onClick={() => setContrast(!contrast)}><Contrast size={19} /></button><button aria-label="Reducir tamaño de letra" onClick={() => setScale(Math.max(.88, +(scale - .06).toFixed(2)))}><Minus size={20} /></button><button aria-label="Aumentar tamaño de letra" onClick={() => setScale(Math.min(1.28, +(scale + .06).toFixed(2)))}><Plus size={20} /></button></aside>
      {metadata && <div className="modal-backdrop" onMouseDown={() => setMetadata(false)}><section className="metadata-modal" role="dialog" aria-modal="true" aria-labelledby="metadata-title" onMouseDown={(event) => event.stopPropagation()}><div className="modal-header"><div><p className="card-kicker">Información del conjunto de datos</p><h2 id="metadata-title">Metadatos de la aplicación</h2></div><button className="icon-button" onClick={() => setMetadata(false)} aria-label="Cerrar metadatos"><X /></button></div><div className="metadata-grid metadata-catalog"><section className="metadata-wide metadata-section"><div className="metadata-section-title"><span>01</span><div><p>Metadatos de Identificación del Conjunto de Datos</p><h3>Batería de indicadores DANE · Putumayo</h3></div></div><div className="metadata-definition-grid"><div><span>Entidad responsable</span><strong>Gobernación del Putumayo · Secretaría de Planeación · Área de Estadística</strong></div><div><span>Propósito</span><strong>Consulta territorial, comparación y descarga de resultados oficiales.</strong></div><div><span>Cobertura territorial</span><strong>Departamento del Putumayo y municipios con información reportada.</strong></div><div><span>Cobertura temporal</span><strong>{metadataSummary.periods}</strong></div><div><span>Fuente de origen</span><strong>{metadataSummary.sources} fuentes de información vinculadas a los registros.</strong></div><div><span>Última fecha de consulta</span><strong>{formatDate(metadataSummary.dates[0] || "")}</strong></div></div></section><section className="metadata-wide metadata-section"><div className="metadata-section-title"><span>02</span><div><p>Metadatos de Estructura del Archivo</p><h3>Composición técnica de la batería consultada</h3></div></div><div className="metadata-definition-grid"><div><span>Base de datos</span><strong>{metadataStructure.mainSheet} · {metadataStructure.mainRecords} registros.</strong></div><div><span>Fichas técnicas</span><strong>{metadataStructure.technicalSheets} fichas técnicas vinculadas.</strong></div><div><span>Unidades de observación</span><strong>Indicador, vigencia, resultado y dimensiones de consulta disponibles.</strong></div><div><span>Modelo territorial</span><strong>Departamento Putumayo, municipios y dimensiones territoriales reportadas para la consulta.</strong></div></div></section><section className="metadata-wide metadata-section metadata-indicators"><div className="metadata-section-title"><span>03</span><div><p>Metadatos Descriptivos de los Indicadores</p><h3>Clasificación, medición y reglas de consulta</h3></div></div><div className="metadata-summary-grid"><div><b>{metadataSummary.records}</b><span>Registros disponibles</span></div><div><b>{metadataSummary.indicators}</b><span>Indicadores</span></div><div><b>{metadataSummary.themes}</b><span>Áreas temáticas</span></div><div><b>{metadataSummary.subthemes}</b><span>Temas</span></div><div><b>{metadataSummary.territories}</b><span>Territorios reportados</span></div><div><b>{metadataSummary.departments}</b><span>Departamentos</span></div><div><b>{metadataSummary.municipalities}</b><span>Municipios disponibles</span></div><div><b>{metadataSummary.divipola}</b><span>Códigos DIVIPOLA</span></div><div><b>{metadataSummary.units}</b><span>Unidades de medida</span></div><div><b>{metadataSummary.odsRecords}</b><span>Registros con ODS o Meta ODS</span></div><div><b>{metadataSummary.sources}</b><span>Fuentes de información</span></div></div></section><section className="metadata-wide metadata-section metadata-quality"><div className="metadata-section-title"><span>04</span><div><p>Metadatos de Calidad (Inferidos)</p><h3>Controles visibles sobre disponibilidad y consistencia</h3></div></div><div className="metadata-definition-grid"><div><span>Resultados numéricos</span><strong>{metadataSummary.numericResults} registros con resultado interpretable.</strong></div><div><span>Enlaces de fuente</span><strong>{metadataSummary.sourcesWithLink} registros con enlace web asociado.</strong></div><div><span>Regla territorial</span><strong>Los resultados se presentan de acuerdo con el territorio y la desagregación seleccionados.</strong></div><div><span>Regla de vigencia</span><strong>La consulta inicia en la vigencia más reciente disponible del indicador elegido.</strong></div><div><span>Precaución metodológica</span><strong>Las comparaciones deben conservar unidad, territorio y desagregación compatibles.</strong></div></div></section><section className="metadata-wide metadata-columns"><span>Diccionario de campos consultables</span><strong>Dimensiones, identificación, resultados y descarga</strong><p>Las definiciones se organizan por las dimensiones visibles en filtros, resultados, ficha técnica y descarga.</p>{metadataGroups.map((group) => <section key={group.title}><h3>{group.title}</h3><div className="metadata-column-list">{group.fields.map((field) => <div key={field}><span>{labels[field] ?? field}</span><strong>{metadataFieldNotes[field] ?? "Campo disponible para consulta."}</strong></div>)}</div></section>)}</section>{activeQuery && <div className="metadata-wide metadata-current"><span>Consulta actual</span><strong>{optionLabel("INDICADOR", code)} · {filters.VIGENCIA}</strong><p>{valueFor(indicator, "AREA TEMATICA", specialIndex)} · {valueFor(indicator, "TEMA", specialIndex)} · {unit}.</p><p>Filtro territorial y de desagregación: {resultFilterContext}. Opciones disponibles: {visibleDimensions.map((field) => field.label).join(" · ") || "No reportadas para esta consulta"}.</p></div>}</div></section></div>}
    </div>
  );
}
