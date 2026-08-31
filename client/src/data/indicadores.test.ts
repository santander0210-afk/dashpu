import { describe, expect, it } from "vitest";
import embeddedJson from "./indicadores.json?raw";

type EmbeddedData = {
  rows: Array<Record<string, unknown>>;
  technicalProfiles: Array<{ code: string }>;
  workbookStructure: Array<{ name: string }>;
};

describe("batería DANE embebida", () => {
  it("conserva la cobertura de la base activa y sus fichas técnicas", () => {
    const data = JSON.parse(embeddedJson) as EmbeddedData;
    const indicators = new Set(data.rows.map((row) => String(row.ID_INDICADOR ?? "")).filter(Boolean));

    expect(data.rows).toHaveLength(1503);
    expect(indicators).toHaveLength(77);
    expect(data.technicalProfiles).toHaveLength(77);
    expect(data.workbookStructure).toHaveLength(78);
    expect(data.rows[0]).toHaveProperty("ID_INDICADOR");
  });
});
