/**
 * 文件名: textUpdate.test.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 描述: textUpdate 工具的单元测试 | textUpdate tool unit tests
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { OfficeMockObject } from 'office-addin-mock';
import { updateTextBox, updateTextBoxes, getTextBoxStyle } from '../../../src/ppt-tools';

type MockShape = {
  id: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  textFrame: {
    textRange: {
      text: string;
      font: {
        name: string;
        size: number;
        color: string;
        bold: boolean;
        italic: boolean;
        underline: string;
        load: () => void;
      };
      paragraphFormat: {
        horizontalAlignment: string;
        load: () => void;
      };
      load: () => void;
    };
    verticalAlignment: string;
    load: () => void;
  };
  fill: {
    type: string;
    foregroundColor: string;
    setSolidColor: (color: string) => void;
    load: () => void;
  };
  load: () => void;
};

type MockData = {
  context: {
    presentation: {
      getSelectedSlides: () => {
        getItemAt: () => {
          shapes: {
            items: MockShape[];
            load: () => void;
          };
        };
      };
    };
    sync: () => Promise<void>;
  };
  run: (callback: (context: MockData['context']) => Promise<void>) => Promise<void>;
  _getShape: (id: string) => MockShape | undefined;
};

// 创建 mock 形状对象
const createMockShape = (
  id: string,
  type: string = 'TextBox',
  initialData?: Partial<MockShape>
): MockShape => {
  const shape: MockShape = {
    id,
    type,
    left: initialData?.left ?? 100,
    top: initialData?.top ?? 100,
    width: initialData?.width ?? 300,
    height: initialData?.height ?? 100,
    textFrame: {
      textRange: {
        text: initialData?.textFrame?.textRange?.text ?? 'Initial Text',
        font: {
          name: 'Arial',
          size: 18,
          color: '#000000',
          bold: false,
          italic: false,
          underline: 'None',
          load: function() {},
        },
        paragraphFormat: {
          horizontalAlignment: 'Left',
          load: function() {},
        },
        load: function() {},
      },
      verticalAlignment: 'Top',
      load: function() {},
    },
    fill: {
      type: 'Solid',
      foregroundColor: '#FFFFFF',
      setSolidColor: function(color: string) {
        this.foregroundColor = color;
      },
      load: function() {},
    },
    load: function() {},
  };
  return shape;
};

// 创建 mock PowerPoint 数据
const createMockData = (shapes: MockShape[]): MockData => {
  return {
    context: {
      presentation: {
        getSelectedSlides: () => ({
          getItemAt: () => ({
            shapes: {
              items: shapes,
              load: function() {},
            },
            load: function() {}, // 支持 slide.load("shapes")
          }),
        }),
      },
      sync: async () => {},
    },
    run: async function(callback: (context: MockData['context']) => Promise<void>) {
      await callback(this.context);
    },
    _getShape: (id: string) => shapes.find((s) => s.id === id),
  };
};

describe('textUpdate 工具测试 | textUpdate Tool Tests', () => {
  beforeEach(() => {
    // 重置 global.PowerPoint
    delete (global as any).PowerPoint;
  });

  describe('updateTextBox - 基础功能测试 | updateTextBox - Basic Functionality Tests', () => {
    it('应该在未提供元素ID时返回错误 | should return error when elementId is not provided', async () => {
      const result = await updateTextBox({ elementId: '' });
      
      expect(result.success).toBe(false);
      expect(result.message).toBe('元素ID不能为空');
    });

    it('应该能够更新文本内容 | should be able to update text content', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        text: 'Updated Text',
      });

      expect(result.success).toBe(true);
      expect(result.message).toBe('文本框更新成功');
      expect(result.elementId).toBe('shape1');
      expect(shape.textFrame.textRange.text).toBe('Updated Text');
    });

    it('应该能够更新字体大小 | should be able to update font size', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        fontSize: 24,
      });

      expect(result.success).toBe(true);
      expect(shape.textFrame.textRange.font.size).toBe(24);
    });

    it('应该能够更新字体名称 | should be able to update font name', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        fontName: 'Times New Roman',
      });

      expect(result.success).toBe(true);
      expect(shape.textFrame.textRange.font.name).toBe('Times New Roman');
    });

    it('应该能够更新字体颜色 | should be able to update font color', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        fontColor: '#FF0000',
      });

      expect(result.success).toBe(true);
      expect(shape.textFrame.textRange.font.color).toBe('#FF0000');
    });

    it('应该能够更新字体样式（加粗、斜体、下划线）| should be able to update font styles (bold, italic, underline)', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        bold: true,
        italic: true,
        underline: true,
      });

      expect(result.success).toBe(true);
      expect(shape.textFrame.textRange.font.bold).toBe(true);
      expect(shape.textFrame.textRange.font.italic).toBe(true);
      expect(shape.textFrame.textRange.font.underline).toBe('Single');
    });

    it('应该能够取消下划线 | should be able to remove underline', async () => {
      const shape = createMockShape('shape1');
      shape.textFrame.textRange.font.underline = 'Single';
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        underline: false,
      });

      expect(result.success).toBe(true);
      expect(shape.textFrame.textRange.font.underline).toBe('None');
    });

    it('应该能够更新水平对齐方式 | should be able to update horizontal alignment', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        horizontalAlignment: 'Center',
      });

      expect(result.success).toBe(true);
      expect(shape.textFrame.textRange.paragraphFormat.horizontalAlignment).toBe('Center');
    });

    it('应该能够更新垂直对齐方式 | should be able to update vertical alignment', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        verticalAlignment: 'Middle',
      });

      expect(result.success).toBe(true);
      expect(shape.textFrame.verticalAlignment).toBe('Middle');
    });

    it('应该能够更新背景颜色 | should be able to update background color', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        backgroundColor: '#00FF00',
      });

      expect(result.success).toBe(true);
      expect(shape.fill.foregroundColor).toBe('#00FF00');
    });

    it('应该能够更新位置和尺寸 | should be able to update position and size', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        left: 200,
        top: 150,
        width: 400,
        height: 200,
      });

      expect(result.success).toBe(true);
      expect(shape.left).toBe(200);
      expect(shape.top).toBe(150);
      expect(shape.width).toBe(400);
      expect(shape.height).toBe(200);
    });

    it('应该能够同时更新多个属性 | should be able to update multiple properties at once', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        text: 'Complete Update',
        fontSize: 20,
        fontName: 'Calibri',
        fontColor: '#0000FF',
        bold: true,
        italic: false,
        underline: true,
        horizontalAlignment: 'Right',
        verticalAlignment: 'Bottom',
        backgroundColor: '#FFFF00',
        left: 50,
        top: 50,
        width: 500,
        height: 150,
      });

      expect(result.success).toBe(true);
      expect(shape.textFrame.textRange.text).toBe('Complete Update');
      expect(shape.textFrame.textRange.font.size).toBe(20);
      expect(shape.textFrame.textRange.font.name).toBe('Calibri');
      expect(shape.textFrame.textRange.font.color).toBe('#0000FF');
      expect(shape.textFrame.textRange.font.bold).toBe(true);
      expect(shape.textFrame.textRange.font.italic).toBe(false);
      expect(shape.textFrame.textRange.font.underline).toBe('Single');
      expect(shape.textFrame.textRange.paragraphFormat.horizontalAlignment).toBe('Right');
      expect(shape.textFrame.verticalAlignment).toBe('Bottom');
      expect(shape.fill.foregroundColor).toBe('#FFFF00');
      expect(shape.left).toBe(50);
      expect(shape.top).toBe(50);
      expect(shape.width).toBe(500);
      expect(shape.height).toBe(150);
    });
  });

  describe('updateTextBox - 错误处理测试 | updateTextBox - Error Handling Tests', () => {
    it('应该在找不到元素时返回错误 | should return error when element is not found', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'nonexistent',
        text: 'Test',
      });

      expect(result.success).toBe(false);
      expect(result.message).toContain('未找到ID为 nonexistent 的元素');
    });

    it('应该在元素类型不支持时返回错误 | should return error when element type is not supported', async () => {
      const shape = createMockShape('shape1', 'Picture');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        text: 'Test',
      });

      expect(result.success).toBe(false);
      expect(result.message).toContain('元素类型 Picture 不支持文本编辑');
    });

    it('应该支持 TextBox 类型 | should support TextBox type', async () => {
      const shape = createMockShape('shape1', 'TextBox');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        text: 'Test',
      });

      expect(result.success).toBe(true);
    });

    it('应该支持 Placeholder 类型 | should support Placeholder type', async () => {
      const shape = createMockShape('shape1', 'Placeholder');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        text: 'Test',
      });

      expect(result.success).toBe(true);
    });

    it('应该支持 GeometricShape 类型 | should support GeometricShape type', async () => {
      const shape = createMockShape('shape1', 'GeometricShape');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        text: 'Test',
      });

      expect(result.success).toBe(true);
    });
  });

  describe('updateTextBoxes - 批量更新测试 | updateTextBoxes - Batch Update Tests', () => {
    it('应该能够批量更新多个文本框 | should be able to update multiple text boxes', async () => {
      const shape1 = createMockShape('shape1');
      const shape2 = createMockShape('shape2');
      const mockData = createMockData([shape1, shape2]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const results = await updateTextBoxes([
        { elementId: 'shape1', text: 'Text 1' },
        { elementId: 'shape2', text: 'Text 2' },
      ]);

      expect(results).toHaveLength(2);
      expect(results[0].success).toBe(true);
      expect(results[1].success).toBe(true);
      expect(shape1.textFrame.textRange.text).toBe('Text 1');
      expect(shape2.textFrame.textRange.text).toBe('Text 2');
    });

    it('应该返回空数组当没有更新项时 | should return empty array when no updates', async () => {
      const mockData = createMockData([]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const results = await updateTextBoxes([]);

      expect(results).toHaveLength(0);
    });

    it('应该继续处理其他项即使某项失败 | should continue processing other items even if one fails', async () => {
      const shape1 = createMockShape('shape1');
      const mockData = createMockData([shape1]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const results = await updateTextBoxes([
        { elementId: 'shape1', text: 'Success' },
        { elementId: 'nonexistent', text: 'Fail' },
        { elementId: 'shape1', fontSize: 30 },
      ]);

      expect(results).toHaveLength(3);
      expect(results[0].success).toBe(true);
      expect(results[1].success).toBe(false);
      expect(results[2].success).toBe(true);
    });
  });

  describe('getTextBoxStyle - 获取样式测试 | getTextBoxStyle - Get Style Tests', () => {
    it('应该在未提供元素ID时返回null | should return null when elementId is not provided', async () => {
      const result = await getTextBoxStyle('');
      expect(result).toBeNull();
    });

    it('应该能够获取文本框的样式信息 | should be able to get text box style information', async () => {
      const shape = createMockShape('shape1');
      shape.textFrame.textRange.text = 'Sample Text';
      shape.textFrame.textRange.font.size = 22;
      shape.textFrame.textRange.font.name = 'Verdana';
      shape.textFrame.textRange.font.color = '#FF00FF';
      shape.textFrame.textRange.font.bold = true;
      shape.textFrame.textRange.font.italic = true;
      shape.textFrame.textRange.font.underline = 'Single';
      shape.textFrame.textRange.paragraphFormat.horizontalAlignment = 'Center';
      shape.textFrame.verticalAlignment = 'Middle';
      shape.fill.foregroundColor = '#00FFFF';
      shape.left = 150;
      shape.top = 200;
      shape.width = 350;
      shape.height = 120;

      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await getTextBoxStyle('shape1');

      expect(result).not.toBeNull();
      expect(result?.elementId).toBe('shape1');
      expect(result?.text).toBe('Sample Text');
      expect(result?.fontSize).toBe(22);
      expect(result?.fontName).toBe('Verdana');
      expect(result?.fontColor).toBe('#FF00FF');
      expect(result?.bold).toBe(true);
      expect(result?.italic).toBe(true);
      expect(result?.underline).toBe(true);
      expect(result?.horizontalAlignment).toBe('Center');
      expect(result?.verticalAlignment).toBe('Middle');
      expect(result?.backgroundColor).toBe('#00FFFF');
      expect(result?.left).toBe(150);
      expect(result?.top).toBe(200);
      expect(result?.width).toBe(350);
      expect(result?.height).toBe(120);
    });

    it('应该正确处理下划线为None的情况 | should correctly handle underline None', async () => {
      const shape = createMockShape('shape1');
      shape.textFrame.textRange.font.underline = 'None';
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await getTextBoxStyle('shape1');

      expect(result?.underline).toBe(false);
    });

    it('应该在找不到元素时返回null | should return null when element is not found', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await getTextBoxStyle('nonexistent');

      expect(result).toBeNull();
    });

    it('应该处理非Solid填充类型 | should handle non-Solid fill type', async () => {
      const shape = createMockShape('shape1');
      shape.fill.type = 'Gradient';
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await getTextBoxStyle('shape1');

      expect(result?.backgroundColor).toBeUndefined();
    });
  });

  describe('边界情况测试 | Edge Case Tests', () => {
    it('应该能够处理空文本 | should be able to handle empty text', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        text: '',
      });

      expect(result.success).toBe(true);
      expect(shape.textFrame.textRange.text).toBe('');
    });

    it('应该能够处理特殊字符 | should be able to handle special characters', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const specialText = '特殊字符 !@#$%^&*() 测试\n换行符';
      const result = await updateTextBox({
        elementId: 'shape1',
        text: specialText,
      });

      expect(result.success).toBe(true);
      expect(shape.textFrame.textRange.text).toBe(specialText);
    });

    it('应该能够处理零值坐标和尺寸 | should be able to handle zero coordinates and size', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        left: 0,
        top: 0,
        width: 0,
        height: 0,
      });

      expect(result.success).toBe(true);
      expect(shape.left).toBe(0);
      expect(shape.top).toBe(0);
      expect(shape.width).toBe(0);
      expect(shape.height).toBe(0);
    });

    it('应该能够处理负值坐标 | should be able to handle negative coordinates', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        left: -10,
        top: -20,
      });

      expect(result.success).toBe(true);
      expect(shape.left).toBe(-10);
      expect(shape.top).toBe(-20);
    });

    it('应该能够处理极小字号 | should be able to handle very small font size', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        fontSize: 1,
      });

      expect(result.success).toBe(true);
      expect(shape.textFrame.textRange.font.size).toBe(1);
    });

    it('应该能够处理极大字号 | should be able to handle very large font size', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const result = await updateTextBox({
        elementId: 'shape1',
        fontSize: 200,
      });

      expect(result.success).toBe(true);
      expect(shape.textFrame.textRange.font.size).toBe(200);
    });

    it('应该能够处理所有对齐方式 | should be able to handle all alignment options', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const alignments: Array<'Left' | 'Center' | 'Right' | 'Justify' | 'Distributed'> = [
        'Left',
        'Center',
        'Right',
        'Justify',
        'Distributed',
      ];

      for (const alignment of alignments) {
        const result = await updateTextBox({
          elementId: 'shape1',
          horizontalAlignment: alignment,
        });

        expect(result.success).toBe(true);
        expect(shape.textFrame.textRange.paragraphFormat.horizontalAlignment).toBe(alignment);
      }
    });

    it('应该能够处理所有垂直对齐方式 | should be able to handle all vertical alignment options', async () => {
      const shape = createMockShape('shape1');
      const mockData = createMockData([shape]);
      (global as any).PowerPoint = new OfficeMockObject(mockData);

      const verticalAlignments: Array<'Top' | 'Middle' | 'Bottom'> = ['Top', 'Middle', 'Bottom'];

      for (const alignment of verticalAlignments) {
        const result = await updateTextBox({
          elementId: 'shape1',
          verticalAlignment: alignment,
        });

        expect(result.success).toBe(true);
        expect(shape.textFrame.verticalAlignment).toBe(alignment);
      }
    });
  });
});
