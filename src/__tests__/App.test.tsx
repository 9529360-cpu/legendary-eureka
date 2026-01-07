import React from "react";
import { render, screen, waitFor, act } from "@testing-library/react";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import App from "../taskpane/components/App";
import { AI_BACKEND_URL } from "../config";

// 错误边界组件
class ErrorBoundary extends React.Component<
  { children: React.ReactNode },
  { hasError: boolean; error: Error | null }
> {
  constructor(props: { children: React.ReactNode }) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: Error) {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo) {
    console.error("Error caught by boundary:", error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div data-testid="error-boundary">
          <h2>组件渲染出错</h2>
          <p>错误: {this.state.error?.message}</p>
          <pre>{this.state.error?.stack}</pre>
        </div>
      );
    }

    return this.props.children;
  }
}

// 包装组件以提供必要的上下文 - 即使App组件内部包含FluentProvider，测试环境也需要外层包装
const WrappedApp = () => (
  <FluentProvider theme={webLightTheme}>
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  </FluentProvider>
);

describe("App Component", () => {
  beforeEach(() => {
    // 重置所有模拟
    jest.clearAllMocks();

    // 设置fetch模拟 - 使用立即解析的Promise，避免异步延迟
    (global.fetch as jest.Mock).mockImplementation(() =>
      Promise.resolve({
        ok: true,
        json: () =>
          Promise.resolve({
            configured: false,
            isValid: false,
            lastUpdated: null,
            maskedKey: null,
          }),
      })
    );

    // 重置localStorage模拟
    (global.localStorage.getItem as jest.Mock).mockReturnValue(null);
    (global.localStorage.setItem as jest.Mock).mockImplementation(() => {});
    (global.localStorage.removeItem as jest.Mock).mockImplementation(() => {});
    (global.localStorage.clear as jest.Mock).mockImplementation(() => {});
  });

  test("renders without crashing", async () => {
    await act(async () => {
      render(<WrappedApp />);
    });

    // 直接检查容器是否存在
    const container = screen.queryByTestId("copilot-container");
    console.log("Container found:", container);
    console.log("Rendered HTML:", document.body.innerHTML);

    // 即使有异步更新，容器应该立即存在
    expect(container).toBeInTheDocument();
  });

  test("contains main container element", async () => {
    await act(async () => {
      render(<WrappedApp />);
    });

    // 等待组件渲染完成
    await waitFor(
      () => {
        const container = screen.getByTestId("copilot-container");
        expect(container).toBeInTheDocument();
      },
      { timeout: 5000 }
    );
  });

  test("renders welcome message", async () => {
    await act(async () => {
      render(<WrappedApp />);
    });

    // 等待组件渲染完成
    await waitFor(
      () => {
        const welcomeText = screen.getByText(/Excel 智能助手/i);
        expect(welcomeText).toBeInTheDocument();
      },
      { timeout: 5000 }
    );
  });

  test("renders header with title", async () => {
    await act(async () => {
      render(<WrappedApp />);
    });

    // 等待组件渲染完成
    await waitFor(
      () => {
        const title = screen.getByText("Excel 智能助手");
        expect(title).toBeInTheDocument();
      },
      { timeout: 5000 }
    );
  });

  test("renders input area", async () => {
    await act(async () => {
      render(<WrappedApp />);
    });

    // 等待组件渲染完成
    await waitFor(
      () => {
        const input = screen.getByPlaceholderText(/输入自然语言指令/i);
        expect(input).toBeInTheDocument();
      },
      { timeout: 5000 }
    );
  });

  test("renders quick actions buttons", async () => {
    await act(async () => {
      render(<WrappedApp />);
    });

    // 等待组件渲染完成
    await waitFor(
      () => {
        const createTableBtn = screen.getByText("创建表格");
        const generateChartBtn = screen.getByText("生成图表");
        const formatBtn = screen.getByText("格式化");

        expect(createTableBtn).toBeInTheDocument();
        expect(generateChartBtn).toBeInTheDocument();
        expect(formatBtn).toBeInTheDocument();
      },
      { timeout: 5000 }
    );
  });
});

describe("Excel API Functions", () => {
  test("Excel.run is defined", () => {
    expect((global as any).Excel.run).toBeDefined();
  });

  test("Office.run is defined", () => {
    expect((global as any).Office.run).toBeDefined();
  });
});

describe("Local Storage Functions", () => {
  test("localStorage is available", () => {
    expect(global.localStorage).toBeDefined();
  });

  test("localStorage methods are mocked", () => {
    const key = "testKey";
    const value = "testValue";

    // 测试setItem被调用
    global.localStorage.setItem(key, value);
    expect(global.localStorage.setItem).toHaveBeenCalledWith(key, value);

    // 测试getItem被调用
    global.localStorage.getItem(key);
    expect(global.localStorage.getItem).toHaveBeenCalledWith(key);

    // 测试removeItem被调用
    global.localStorage.removeItem(key);
    expect(global.localStorage.removeItem).toHaveBeenCalledWith(key);

    // 测试clear被调用
    global.localStorage.clear();
    expect(global.localStorage.clear).toHaveBeenCalled();
  });
});

describe("API Key Management", () => {
  beforeEach(() => {
    // 在每个测试前重置fetch模拟
    (global.fetch as jest.Mock).mockClear();
  });

  test("fetches API key status on mount", async () => {
    // 创建一个立即解析的Promise来模拟fetch响应
    const mockResponse = {
      ok: true,
      json: () =>
        Promise.resolve({
          configured: false,
          isValid: false,
          lastUpdated: null,
          maskedKey: null,
        }),
    };

    (global.fetch as jest.Mock).mockResolvedValue(mockResponse);

    // 使用act包装整个异步渲染过程
    await act(async () => {
      render(<WrappedApp />);
    });

    // 等待fetch被调用
    await waitFor(
      () => {
        expect(global.fetch).toHaveBeenCalledWith(`${AI_BACKEND_URL}/api/config/status`);
      },
      { timeout: 5000 }
    );
  });

  test("handles API key status with cached data", async () => {
    // 模拟localStorage中有缓存数据
    const cachedStatus = {
      status: {
        configured: true,
        isValid: true,
        lastUpdated: new Date().toISOString(),
        maskedKey: "sk-63af...0fdb",
      },
      cachedAt: new Date().toISOString(),
    };

    (global.localStorage.getItem as jest.Mock).mockReturnValue(JSON.stringify(cachedStatus));

    // 模拟fetch响应
    (global.fetch as jest.Mock).mockResolvedValue({
      ok: true,
      json: () =>
        Promise.resolve({
          configured: true,
          isValid: true,
          lastUpdated: new Date().toISOString(),
          maskedKey: "sk-63af...0fdb",
        }),
    });

    await act(async () => {
      render(<WrappedApp />);
    });

    // 等待fetch被调用
    await waitFor(
      () => {
        expect(global.fetch).toHaveBeenCalledWith(`${AI_BACKEND_URL}/api/config/status`);
      },
      { timeout: 5000 }
    );
  });

  test("handles API key status fetch error", async () => {
    // 模拟fetch失败
    (global.fetch as jest.Mock).mockRejectedValue(new Error("Network error"));

    await act(async () => {
      render(<WrappedApp />);
    });

    // 等待组件渲染完成（即使有错误）
    await waitFor(
      () => {
        const container = screen.getByTestId("copilot-container");
        expect(container).toBeInTheDocument();
      },
      { timeout: 5000 }
    );
  });
});
