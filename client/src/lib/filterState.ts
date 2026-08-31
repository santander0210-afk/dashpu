export function clearDimensionSelections(
  current: Record<string, string>,
  dimensionKeys: readonly string[],
  departmentKey: string,
  allValue: string,
  departmentValue = "Putumayo",
) {
  return {
    ...current,
    ...Object.fromEntries(
      dimensionKeys.map((key) => [key, key === departmentKey ? departmentValue : allValue]),
    ),
  };
}
