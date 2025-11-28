/**
 * 文件名: test-utils.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/28
 * 最后修改日期: 2025/11/28
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: @testing-library/react, @fluentui/react-components
 * 描述: 测试工具函数 | Test utility functions
 */

import * as React from 'react';
import { ReactElement } from 'react';
import { render, RenderOptions } from '@testing-library/react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { vi } from 'vitest';

/**
 * 自定义渲染函数，包装 FluentUI Provider
 * Custom render function with FluentUI Provider wrapper
 */
interface CustomRenderOptions extends Omit<RenderOptions, 'wrapper'> {
  theme?: typeof webLightTheme;
}

export function renderWithProviders(
  ui: ReactElement,
  options?: CustomRenderOptions
) {
  const { theme = webLightTheme, ...renderOptions } = options || {};

  function Wrapper({ children }: { children: React.ReactNode }) {
    return <FluentProvider theme={theme}>{children}</FluentProvider>;
  }

  return render(ui, { wrapper: Wrapper, ...renderOptions });
}

/**
 * 创建模拟的 Office 上下文
 * Create mock Office context
 */
export function createMockOfficeContext() {
  return {
    presentation: {
      slides: {
        getItemAt: vi.fn().mockReturnValue({
          shapes: {
            addTextBox: vi.fn().mockReturnValue({
              textFrame: {
                textRange: {
                  text: '',
                },
              },
              left: 0,
              top: 0,
              load: vi.fn(),
            }),
          },
          load: vi.fn(),
        }),
        add: vi.fn(),
      },
      load: vi.fn(),
    },
    sync: vi.fn().mockResolvedValue(undefined),
  };
}

/**
 * 等待异步操作完成
 * Wait for async operations to complete
 */
export const waitForAsync = () => new Promise((resolve) => setTimeout(resolve, 0));

/**
 * 模拟 PowerPoint.run 调用
 * Mock PowerPoint.run call
 */
export function mockPowerPointRun(mockContext?: any) {
  const context = mockContext || createMockOfficeContext();
  return vi.fn((callback) => Promise.resolve(callback(context)));
}

// 重新导出所有 testing-library 工具 | Re-export all testing-library utilities
export * from '@testing-library/react';
export { default as userEvent } from '@testing-library/user-event';
