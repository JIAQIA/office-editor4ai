/**
 * 文件名: textInsertion.test.ts
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 描述: textInsertion 工具的单元测试
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { OfficeMockObject } from 'office-addin-mock';
import { insertText, insertTextToSlide } from '../../../src/ppt-tools';

type MockTextBox = {
  fill: {
    color: string;
    setSolidColor: (color: string) => void;
  };
  lineFormat: {
    color: string;
    weight: number;
    dashStyle: string;
  };
  _text?: string;
  _options?: unknown;
};

type MockData = {
  context: {
    presentation: {
      getSelectedSlides: () => {
        getItemAt: () => {
          shapes: {
            addTextBox: (text: string, options?: unknown) => MockTextBox;
          };
        };
      };
    };
  };
  run: (callback: (context: MockData['context']) => Promise<void>) => Promise<void>;
  _getTextBox: () => MockTextBox;
};

// 创建 mock 文本框对象
const createMockTextBox = (): MockTextBox => ({
  fill: {
    color: '',
    setSolidColor: function(color: string) {
      this.color = color;
    },
  },
  lineFormat: {
    color: '',
    weight: 0,
    dashStyle: '',
  },
});

// 创建 mock PowerPoint 数据
const createMockData = (): MockData => {
  const mockTextBox = createMockTextBox();
  return {
    context: {
      presentation: {
        getSelectedSlides: () => ({
          getItemAt: () => ({
            shapes: {
              addTextBox: (text: string, options?: unknown) => {
                mockTextBox._text = text;
                mockTextBox._options = options;
                return mockTextBox;
              },
            },
          }),
        }),
      },
    },
    run: async function(callback: (context: MockData['context']) => Promise<void>) {
      await callback(this.context);
    },
    _getTextBox: () => mockTextBox,
  };
};

describe('textInsertion 工具测试', () => {
  beforeEach(() => {
    // 重置 global.PowerPoint
    delete (global as any).PowerPoint;
  });

  describe('insertTextToSlide', () => {
    it('应该能够插入带有默认参数的文本框', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertTextToSlide({ text: 'Hello World' });

      const textBox = mockData._getTextBox();
      expect(textBox._text).toBe('Hello World');
      expect(textBox._options).toBeUndefined();
      expect(textBox.fill.color).toBe('white');
      expect(textBox.lineFormat.color).toBe('black');
      expect(textBox.lineFormat.weight).toBe(1);
      expect(textBox.lineFormat.dashStyle).toBe('Solid');
    });

    it('应该能够插入带有指定位置的文本框', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertTextToSlide({
        text: 'Positioned Text',
        left: 100,
        top: 200,
      });

      const textBox = mockData._getTextBox();
      expect(textBox._text).toBe('Positioned Text');
      expect(textBox._options).toEqual({
        left: 100,
        top: 200,
        width: 300,
        height: 100,
      });
    });

    it('应该能够插入带有自定义尺寸的文本框', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertTextToSlide({
        text: 'Custom Size',
        left: 50,
        top: 50,
        width: 400,
        height: 150,
      });

      const textBox = mockData._getTextBox();
      expect(textBox._text).toBe('Custom Size');
      expect(textBox._options).toEqual({
        left: 50,
        top: 50,
        width: 400,
        height: 150,
      });
    });

    it('应该能够插入带有自定义样式的文本框', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertTextToSlide({
        text: 'Styled Text',
        fillColor: 'blue',
        lineColor: 'red',
        lineWeight: 2,
      });

      const textBox = mockData._getTextBox();
      expect(textBox.fill.color).toBe('blue');
      expect(textBox.lineFormat.color).toBe('red');
      expect(textBox.lineFormat.weight).toBe(2);
    });

    it('应该能够插入带有完整配置的文本框', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertTextToSlide({
        text: 'Full Config',
        left: 100,
        top: 200,
        width: 500,
        height: 200,
        fillColor: 'yellow',
        lineColor: 'green',
        lineWeight: 3,
      });

      const textBox = mockData._getTextBox();
      expect(textBox._text).toBe('Full Config');
      expect(textBox._options).toEqual({
        left: 100,
        top: 200,
        width: 500,
        height: 200,
      });
      expect(textBox.fill.color).toBe('yellow');
      expect(textBox.lineFormat.color).toBe('green');
      expect(textBox.lineFormat.weight).toBe(3);
    });

    it('应该在只指定 left 时使用默认位置', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertTextToSlide({
        text: 'Only Left',
        left: 100,
      });

      const textBox = mockData._getTextBox();
      expect(textBox._text).toBe('Only Left');
      // 只有 left 没有 top，应该使用默认位置（不传位置参数）
      expect(textBox._options).toBeUndefined();
    });

    it('应该在只指定 top 时使用默认位置', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertTextToSlide({
        text: 'Only Top',
        top: 200,
      });

      const textBox = mockData._getTextBox();
      expect(textBox._text).toBe('Only Top');
      // 只有 top 没有 left，应该使用默认位置（不传位置参数）
      expect(textBox._options).toBeUndefined();
    });

    it('应该在插入失败时抛出错误', async () => {
      const mockData = {
        context: {
          presentation: {
            getSelectedSlides: function() {
              throw new Error('插入失败');
            },
          },
        },
        run: async function(callback: (context: typeof mockData.context) => Promise<void>) {
          await callback(this.context);
        },
      };

      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await expect(insertTextToSlide({ text: 'Error Test' })).rejects.toThrow();
    });
  });

  describe('insertText', () => {
    it('应该能够插入简单文本', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertText('Simple Text');

      const textBox = mockData._getTextBox();
      expect(textBox._text).toBe('Simple Text');
      expect(textBox._options).toBeUndefined();
    });

    it('应该能够插入带有位置的文本', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertText('Positioned', 150, 250);

      const textBox = mockData._getTextBox();
      expect(textBox._text).toBe('Positioned');
      expect(textBox._options).toEqual({
        left: 150,
        top: 250,
        width: 300,
        height: 100,
      });
    });

    it('应该能够只传入文本内容', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertText('Text Only');

      const textBox = mockData._getTextBox();
      expect(textBox._text).toBe('Text Only');
    });

    it('应该正确调用 insertTextToSlide', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const text = 'Test';
      const left = 100;
      const top = 200;

      await insertText(text, left, top);

      const textBox = mockData._getTextBox();
      expect(textBox._text).toBe(text);
      expect(textBox._options).toEqual({
        left,
        top,
        width: 300,
        height: 100,
      });
    });
  });

  describe('边界情况测试', () => {
    it('应该能够插入空字符串', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertText('');

      const textBox = mockData._getTextBox();
      expect(textBox._text).toBe('');
    });

    it('应该能够插入包含特殊字符的文本', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const specialText = '特殊字符 !@#$%^&*() 测试';
      await insertText(specialText);

      const textBox = mockData._getTextBox();
      expect(textBox._text).toBe(specialText);
    });

    it('应该能够插入多行文本', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const multilineText = '第一行\n第二行\n第三行';
      await insertText(multilineText);

      const textBox = mockData._getTextBox();
      expect(textBox._text).toBe(multilineText);
    });

    it('应该能够处理零坐标', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertText('Zero Position', 0, 0);

      const textBox = mockData._getTextBox();
      expect(textBox._text).toBe('Zero Position');
      expect(textBox._options).toEqual({
        left: 0,
        top: 0,
        width: 300,
        height: 100,
      });
    });

    it('应该能够处理负坐标', async () => {
      const mockData = createMockData();
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      await insertText('Negative Position', -10, -20);

      const textBox = mockData._getTextBox();
      expect(textBox._text).toBe('Negative Position');
      expect(textBox._options).toEqual({
        left: -10,
        top: -20,
        width: 300,
        height: 100,
      });
    });
  });
});
