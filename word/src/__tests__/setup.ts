import { beforeEach, vi } from "vitest";

export type MockWordContext = ReturnType<typeof createMockWordContext>;
export let mockWordContext: MockWordContext;
export const wordRun = vi.fn();
export const isSetSupported = vi.fn();

const localStorageData = new Map<string, string>();

Object.defineProperty(globalThis, "localStorage", {
  configurable: true,
  value: {
    clear: () => localStorageData.clear(),
    getItem: (key: string) => localStorageData.get(key) ?? null,
    key: (index: number) => [...localStorageData.keys()][index] ?? null,
    get length() {
      return localStorageData.size;
    },
    removeItem: (key: string) => localStorageData.delete(key),
    setItem: (key: string, value: string) => localStorageData.set(key, String(value))
  }
});

export function createMockWordContext() {
  const paragraphItems = [
    { text: "Sample paragraph.", load: vi.fn(), style: "Normal" },
    { text: "Another paragraph.", load: vi.fn(), style: "Normal" },
    { text: "Third paragraph.", load: vi.fn(), style: "Normal" }
  ];
  const insertedRange = { font: { highlightColor: undefined as string | undefined } };
  const selection = {
    text: "selected content here",
    start: 0,
    end: 21,
    insertText: vi.fn().mockReturnValue(insertedRange),
    insertOoxml: vi.fn().mockReturnValue({}),
    insertComment: vi.fn().mockReturnValue({}),
    font: {
      bold: undefined as boolean | undefined,
      italic: undefined as boolean | undefined,
      color: undefined as string | undefined,
      highlightColor: undefined as string | undefined,
      size: undefined as number | undefined
    },
    paragraphs: { items: paragraphItems, load: vi.fn() },
    load: vi.fn()
  };
  const matches = {
    items: [
      { insertText: vi.fn().mockReturnValue({ font: {} }) },
      { insertText: vi.fn().mockReturnValue({ font: {} }) }
    ],
    load: vi.fn()
  };
  const body = {
    text: "Sample paragraph.\nAnother paragraph.\nThird paragraph.",
    paragraphs: { items: paragraphItems, load: vi.fn() },
    insertText: vi.fn().mockReturnValue(insertedRange),
    insertOoxml: vi.fn().mockReturnValue({}),
    search: vi.fn().mockReturnValue(matches),
    load: vi.fn()
  };
  return {
    document: { body, getSelection: vi.fn().mockReturnValue(selection) },
    sync: vi.fn().mockResolvedValue(undefined),
    selection,
    body,
    matches
  };
}

beforeEach(() => {
  vi.clearAllMocks();
  localStorage.clear();
  mockWordContext = createMockWordContext();
  wordRun.mockImplementation((work: (context: MockWordContext) => Promise<unknown>) => work(mockWordContext));
  isSetSupported.mockReturnValue(true);
  (globalThis as unknown as Record<string, unknown>).Word = { run: wordRun };
  (globalThis as unknown as Record<string, unknown>).Office = {
    context: { requirements: { isSetSupported } }
  };
});
