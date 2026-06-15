import { FormulaNode } from './formula.types';

export class DependencyGraph {
  private readonly dependents = new Map<string, Set<string>>();

  addDependency(source: string, dependent: string): void {
    if (!this.dependents.has(source)) {
      this.dependents.set(source, new Set());
    }
    this.dependents.get(source)!.add(dependent);
  }

  getDependents(cellKey: string): string[] {
    return Array.from(this.dependents.get(cellKey) ?? []);
  }

  getAllSources(): string[] {
    return Array.from(this.dependents.keys());
  }
}

export function buildDependencyGraph(nodes: FormulaNode[]): DependencyGraph {
  const graph = new DependencyGraph();

  for (const node of nodes) {
    const nodeKey = `${node.sheetName}!${node.address}`;
    for (const ref of node.formula.cellRefs) {
      const sourceKey = ref.sheet
        ? `${ref.sheet}!${ref.column}${ref.row}`
        : `${node.sheetName}!${ref.column}${ref.row}`;
      graph.addDependency(sourceKey, nodeKey);
    }
    for (const range of node.formula.rangeRefs) {
      const sourceKey = range.sheet
        ? `${range.sheet}!${range.startCol}${range.startRow}:${range.endCol}${range.endRow}`
        : `${node.sheetName}!${range.startCol}${range.startRow}:${range.endCol}${range.endRow}`;
      graph.addDependency(sourceKey, nodeKey);
    }
  }

  return graph;
}
