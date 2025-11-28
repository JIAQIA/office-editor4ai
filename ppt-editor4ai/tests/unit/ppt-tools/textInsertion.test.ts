/**
 * 文件名: textInsertion.test.ts
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 描述: textInsertion 工具的单元测试
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { insertText, insertTextToSlide } from '../../../src/ppt-tools';

// Mock PowerPoint API
const mockTextBox = {
  fill: {
    setSolidColor: vi.fn(),
  },
  lineFormat: {
    color: '',
    weight: 0,
    dashStyle: '',
  },
};

const mockSlide = {
  shapes: {
    addTextBox: vi.fn(() => mockTextBox),
  },
};

const mockContext = {
  presentation: {
    getSelectedSlides: vi.fn(() => ({
      getItemAt: vi.fn(() => mockSlide),
    })),
  },
  sync: vi.fn(),
};

global.PowerPoint = {
  run: vi.fn((callback) => callback(mockContext)),
} as any;

describe('textInsertion 工具测试', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe('insertTextToSlide', () => {
    it('应该能够插入带有默认参数的文本框', async () => {
      await insertTextToSlide({ text: 'Hello World' });

      expect(mockSlide.shapes.addTextBox).toHaveBeenCalledWith('Hello World');
      expect(mockTextBox.fill.setSolidColor).toHaveBeenCalledWith('white');
      expect(mockTextBox.lineFormat.color).toBe('black');
      expect(mockTextBox.lineFormat.weight).toBe(1);
      expect(mockTextBox.lineFormat.dashStyle).toBe('Solid');
      expect(mockContext.sync).toHaveBeenCalled();
    });

    it('应该能够插入带有指定位置的文本框', async () => {
      await insertTextToSlide({
        text: 'Positioned Text',
        left: 100,
        top: 200,
      });

      expect(mockSlide.shapes.addTextBox).toHaveBeenCalledWith('Positioned Text', {
        left: 100,
        top: 200,
        width: 300,
        height: 100,
      });
    });

    it('应该能够插入带有自定义尺寸的文本框', async () => {
      await insertTextToSlide({
        text: 'Custom Size',
        left: 50,
        top: 50,
        width: 400,
        height: 150,
      });

      expect(mockSlide.shapes.addTextBox).toHaveBeenCalledWith('Custom Size', {
        left: 50,
        top: 50,
        width: 400,
        height: 150,
      });
    });

    it('应该能够插入带有自定义样式的文本框', async () => {
      await insertTextToSlide({
        text: 'Styled Text',
        fillColor: 'blue',
        lineColor: 'red',
        lineWeight: 2,
      });

      expect(mockTextBox.fill.setSolidColor).toHaveBeenCalledWith('blue');
      expect(mockTextBox.lineFormat.color).toBe('red');
      expect(mockTextBox.lineFormat.weight).toBe(2);
    });

    it('应该能够插入带有完整配置的文本框', async () => {
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

      expect(mockSlide.shapes.addTextBox).toHaveBeenCalledWith('Full Config', {
        left: 100,
        top: 200,
        width: 500,
        height: 200,
      });
      expect(mockTextBox.fill.setSolidColor).toHaveBeenCalledWith('yellow');
      expect(mockTextBox.lineFormat.color).toBe('green');
      expect(mockTextBox.lineFormat.weight).toBe(3);
    });

    it('应该在只指定 left 时使用默认位置', async () => {
      await insertTextToSlide({
        text: 'Only Left',
        left: 100,
      });

      // 只有 left 没有 top，应该使用默认位置（不传位置参数）
      expect(mockSlide.shapes.addTextBox).toHaveBeenCalledWith('Only Left');
    });

    it('应该在只指定 top 时使用默认位置', async () => {
      await insertTextToSlide({
        text: 'Only Top',
        top: 200,
      });

      // 只有 top 没有 left，应该使用默认位置（不传位置参数）
      expect(mockSlide.shapes.addTextBox).toHaveBeenCalledWith('Only Top');
    });

    it('应该在插入失败时抛出错误', async () => {
      const error = new Error('插入失败');
      mockSlide.shapes.addTextBox.mockImplementationOnce(() => {
        throw error;
      });

      await expect(insertTextToSlide({ text: 'Error Test' })).rejects.toThrow();
    });
  });

  describe('insertText', () => {
    it('应该能够插入简单文本', async () => {
      await insertText('Simple Text');

      expect(mockSlide.shapes.addTextBox).toHaveBeenCalledWith('Simple Text');
      expect(mockContext.sync).toHaveBeenCalled();
    });

    it('应该能够插入带有位置的文本', async () => {
      await insertText('Positioned', 150, 250);

      expect(mockSlide.shapes.addTextBox).toHaveBeenCalledWith('Positioned', {
        left: 150,
        top: 250,
        width: 300,
        height: 100,
      });
    });

    it('应该能够只传入文本内容', async () => {
      await insertText('Text Only');

      expect(mockSlide.shapes.addTextBox).toHaveBeenCalledWith('Text Only');
    });

    it('应该正确调用 insertTextToSlide', async () => {
      const text = 'Test';
      const left = 100;
      const top = 200;

      await insertText(text, left, top);

      expect(mockSlide.shapes.addTextBox).toHaveBeenCalledWith(text, {
        left,
        top,
        width: 300,
        height: 100,
      });
    });
  });

  describe('边界情况测试', () => {
    it('应该能够插入空字符串', async () => {
      await insertText('');

      expect(mockSlide.shapes.addTextBox).toHaveBeenCalledWith('');
    });

    it('应该能够插入包含特殊字符的文本', async () => {
      const specialText = '特殊字符 !@#$%^&*() 测试';
      await insertText(specialText);

      expect(mockSlide.shapes.addTextBox).toHaveBeenCalledWith(specialText);
    });

    it('应该能够插入多行文本', async () => {
      const multilineText = '第一行\n第二行\n第三行';
      await insertText(multilineText);

      expect(mockSlide.shapes.addTextBox).toHaveBeenCalledWith(multilineText);
    });

    it('应该能够处理零坐标', async () => {
      await insertText('Zero Position', 0, 0);

      expect(mockSlide.shapes.addTextBox).toHaveBeenCalledWith('Zero Position', {
        left: 0,
        top: 0,
        width: 300,
        height: 100,
      });
    });

    it('应该能够处理负坐标', async () => {
      await insertText('Negative Position', -10, -20);

      expect(mockSlide.shapes.addTextBox).toHaveBeenCalledWith('Negative Position', {
        left: -10,
        top: -20,
        width: 300,
        height: 100,
      });
    });
  });
});
