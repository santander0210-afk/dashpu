import { describe, expect, it } from "vitest";
import { clearDimensionSelections } from "./filterState";

describe("clearDimensionSelections", () => {
  it("conserva Putumayo y limpia las demás dimensiones", () => {
    const result = clearDimensionSelections(
      {
        "TERRITORIAL DEPARTAMENTO": "Mocoa",
        "TERRITORIAL MUNICIPIO": "Mocoa",
        "ÁREA_GEOGRÁFICA": "Urbano",
        SEXO: "Mujeres",
      },
      ["TERRITORIAL DEPARTAMENTO", "TERRITORIAL MUNICIPIO", "ÁREA_GEOGRÁFICA", "SEXO"],
      "TERRITORIAL DEPARTAMENTO",
      "__TODOS__",
    );

    expect(result).toEqual({
      "TERRITORIAL DEPARTAMENTO": "Putumayo",
      "TERRITORIAL MUNICIPIO": "__TODOS__",
      "ÁREA_GEOGRÁFICA": "__TODOS__",
      SEXO: "__TODOS__",
    });
  });
});
