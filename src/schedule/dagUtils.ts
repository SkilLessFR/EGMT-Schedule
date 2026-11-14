export function encodeStage3DagState(w: number, sh: number, ml: number, sb: number): number {
  return ((w * 4 + sh) * 13 + ml) * 2 + sb;
}

export function encodeGenDagState(w: number, sh: number, ml: number, offS: number, sb: number): number {
  return (((w * 4 + sh) * 13 + ml) * 4 + offS) * 2 + sb;
}

export function canDoShift(dayActMasks: Uint8Array[], eIdx: number, d: number, act: number): boolean {
  return (dayActMasks[eIdx][d] & (1 << act)) !== 0;
}

export function buildDagDayActMasks(
  dags: ReadonlyArray<{
    feasible: boolean;
    layers: ReadonlyArray<ReadonlyArray<ReadonlyArray<{ act: number }>>>;
  }>,
  daysCount: number
): Uint8Array[] {
  return dags.map((dag) => {
    const masks = new Uint8Array(daysCount);
    if (!dag.feasible) return masks;
    for (let d = 0; d < daysCount; d++) {
      let mask = 0;
      const layer = dag.layers[d];
      for (let u = 0; u < layer.length; u++) {
        const edges = layer[u];
        for (let ei = 0; ei < edges.length; ei++) {
          mask |= 1 << edges[ei].act;
        }
      }
      masks[d] = mask;
    }
    return masks;
  });
}

export function isDagPathValid(
  dag: {
    layers: ReadonlyArray<ReadonlyArray<ReadonlyArray<{ act: number; nextNode: number }>>>;
  },
  path: number[],
  daysCount: number
): boolean {
  let u = 0;
  for (let d = 0; d < daysCount; d++) {
    const edges = dag.layers[d][u];
    const act = path[d];
    let found = -1;
    for (let ei = 0; ei < edges.length; ei++) {
      if (edges[ei].act === act) {
        found = edges[ei].nextNode;
        break;
      }
    }
    if (found === -1) return false;
    u = found;
  }
  return true;
}

export function solveDagExtremeCount(
  dag: {
    maxNodes: number;
    layers: ReadonlyArray<ReadonlyArray<ReadonlyArray<{ act: number; nextNode: number }>>>;
  },
  daysCount: number,
  targetAct: number,
  sign: 1 | -1
): number {
  let dpNext = new Float64Array(Math.max(1, dag.maxNodes));
  let dpCur = new Float64Array(Math.max(1, dag.maxNodes));
  for (let d = daysCount - 1; d >= 0; d--) {
    const layer = dag.layers[d];
    const isLast = d === daysCount - 1;
    for (let u = 0; u < layer.length; u++) {
      const edges = layer[u];
      let bestVal = Infinity;
      for (let ei = 0; ei < edges.length; ei++) {
        const edge = edges[ei];
        const val = (edge.act === targetAct ? sign : 0) + (isLast ? 0 : dpNext[edge.nextNode]);
        if (val < bestVal) bestVal = val;
      }
      dpCur[u] = bestVal;
    }
    const tmp = dpNext;
    dpNext = dpCur;
    dpCur = tmp;
  }
  return sign === 1 ? dpNext[0] : -dpNext[0];
}
