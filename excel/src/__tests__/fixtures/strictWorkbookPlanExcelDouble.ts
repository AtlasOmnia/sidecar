// R7-H local strict, stateful Office subset. No read without load + sync;
// writes commit only on sync; foreign sheets/unsupported APIs fail closed.
export class StrictWorkbookPlanExcelDouble {
  readonly sheets = new Map<string, unknown[][]>();
  writes = 0;
  runs = 0;
  violations: string[] = [];
  readonly run = async (callback: (context: unknown) => Promise<unknown>) => {
    this.runs++;
    const pending: (() => void)[] = [];
    const object = (data: Record<string, () => unknown>, methods: Record<string, unknown> = {}, setters: Record<string, (v: unknown) => void> = {}) => {
      const loaded = new Set<string>();
      const snapshot = new Map<string, unknown>();
      return new Proxy({ ...methods, load: (fields: string | string[]) => {
        for (const field of (Array.isArray(fields) ? fields : fields.split(","))) {
          const key = field.split("/")[0];
          if (!(key in data)) throw new Error(`Unsupported load ${field}`);
          pending.push(() => { loaded.add(key); snapshot.set(key, data[key]()); });
        }
      } }, {
        get: (target, key) => {
          if (typeof key !== "string") return undefined;
          if (key in target) return (target as Record<string, unknown>)[key];
          if (!(key in data)) throw new Error(`Unsupported property ${key}`);
          if (!loaded.has(key)) { this.violations.push(key); throw new Error(`read before load+sync: ${key}`); }
          return snapshot.get(key);
        },
        set: (_target, key, value) => {
          if (typeof key !== "string" || !(key in setters)) throw new Error(`Unsupported write ${String(key)}`);
          pending.push(() => { setters[key](value); this.writes++; }); return true;
        },
      });
    };
    const range = (name: string, address: string) => {
      if (address !== "A1") throw new Error(`Unsupported synthetic range ${address}`);
      return object({
        address: () => `${name}!A1`, rowCount: () => 1, columnCount: () => 1,
        values: () => structuredClone(this.sheets.get(name)!),
        formulas: () => [[""]], numberFormat: () => [["General"]], isNullObject: () => false,
      }, {}, { values: v => { this.sheets.set(name, structuredClone(v as unknown[][])); } });
    };
    const sheet = (name: string, nullable = false) => {
      if (!nullable && !this.sheets.has(name)) throw new Error("Foreign worksheet");
      return object({ name: () => name, isNullObject: () => !this.sheets.has(name) }, {
        getRange: (address: string) => range(name, address),
        getUsedRangeOrNullObject: () => range(name, "A1"),
      });
    };
    const names = [...this.sheets.keys()];
    // Nested items/name is loaded as a single host collection operation.
    const children = names.map(n => sheet(n));
    let itemsLoaded = false;
    const finalCollection = {
      getItem: (n: string) => sheet(n), getItemOrNullObject: (n: string) => sheet(n, true),
      getActiveWorksheet: () => sheet(names[0]),
      load: () => {
        for (const child of children) child.load("name");
        pending.push(() => { itemsLoaded = true; });
      },
      get items() { if (!itemsLoaded) throw new Error("items before sync"); return children; },
    };
    const tables = object({ items: () => [] }, {
      getItemOrNullObject: () => object({ isNullObject: () => true }),
    });
    const context = { workbook: { worksheets: finalCollection, tables,
      getSelectedRange: () => range(names[0], "A1") },
      sync: async () => { const batch = pending.splice(0); for (const action of batch) action(); },
    };
    return callback(context);
  };
}
