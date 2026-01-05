// Jest测试配置文件
import "@testing-library/jest-dom";

// 模拟Office.js API
(global as any).Office = {
  context: {
    document: {},
    workbook: {},
  },
  run: jest.fn((callback) => Promise.resolve(callback({}))),
  initialize: jest.fn(),
  onReady: jest.fn(),
  Actions: {
    associate: jest.fn(),
  },
} as any;

// 模拟Excel API
(global as any).Excel = {
  run: jest.fn((callback) => Promise.resolve(callback({}))),
  createWorkbook: jest.fn(),
  load: jest.fn(),
  sync: jest.fn(),
} as any;

// 模拟localStorage
const localStorageMock = {
  getItem: jest.fn(),
  setItem: jest.fn(),
  removeItem: jest.fn(),
  clear: jest.fn(),
  length: 0,
  key: jest.fn(),
};
Object.defineProperty(global, "localStorage", {
  value: localStorageMock,
  writable: true,
});

// 模拟sessionStorage
const sessionStorageMock = {
  getItem: jest.fn(),
  setItem: jest.fn(),
  removeItem: jest.fn(),
  clear: jest.fn(),
  length: 0,
  key: jest.fn(),
};
Object.defineProperty(global, "sessionStorage", {
  value: sessionStorageMock,
  writable: true,
});

// 模拟fetch API
Object.defineProperty(global, "fetch", {
  value: jest.fn(),
  writable: true,
});

// 模拟ResizeObserver
(global as any).ResizeObserver = jest.fn().mockImplementation(() => ({
  observe: jest.fn(),
  unobserve: jest.fn(),
  disconnect: jest.fn(),
}));

// 模拟IntersectionObserver
(global as any).IntersectionObserver = jest.fn().mockImplementation(() => ({
  observe: jest.fn(),
  unobserve: jest.fn(),
  disconnect: jest.fn(),
  takeRecords: jest.fn(),
}));

// 清理所有模拟
afterEach(() => {
  jest.clearAllMocks();
  localStorageMock.clear();
  sessionStorageMock.clear();
});

// 测试超时设置
jest.setTimeout(10000);
