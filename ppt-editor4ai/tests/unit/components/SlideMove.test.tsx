/**
 * 文件名: SlideMove.test.tsx
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 描述: SlideMove 组件的单元测试 | SlideMove component unit tests
 */

import { describe, it, expect, beforeEach, vi } from 'vitest';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import { SlideMove } from '../../../src/taskpane/components/tools/SlideMove';
import * as pptTools from '../../../src/ppt-tools';

// Mock ppt-tools 模块
vi.mock('../../../src/ppt-tools', () => ({
  moveSlide: vi.fn(),
  moveCurrentSlide: vi.fn(),
  swapSlides: vi.fn(),
  getAllSlidesInfo: vi.fn(),
}));

// Mock PowerPoint global
const mockPowerPoint = {
  run: vi.fn(),
};

describe('SlideMove 组件', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    // @ts-expect-error - Mock PowerPoint global
    global.PowerPoint = mockPowerPoint;

    // 默认 mock 返回值
    vi.mocked(pptTools.getAllSlidesInfo).mockResolvedValue([
      { index: 1, id: 'slide-1', title: 'Slide 1' },
      { index: 2, id: 'slide-2', title: 'Slide 2' },
      { index: 3, id: 'slide-3', title: 'Slide 3' },
      { index: 4, id: 'slide-4', title: 'Slide 4' },
      { index: 5, id: 'slide-5', title: 'Slide 5' },
    ]);

    mockPowerPoint.run.mockImplementation(async (callback: any) => {
      const mockContext = {
        presentation: {
          getSelectedSlides: () => ({
            items: [{ id: 'slide-1' }],
            load: vi.fn(),
          }),
          slides: {
            items: [
              { id: 'slide-1' },
              { id: 'slide-2' },
              { id: 'slide-3' },
              { id: 'slide-4' },
              { id: 'slide-5' },
            ],
            load: vi.fn(),
          },
        },
        sync: vi.fn().mockResolvedValue(undefined),
      };

      // 为每个 slide 添加 load 方法
      mockContext.presentation.slides.items.forEach((slide: any) => {
        slide.load = vi.fn();
      });

      await callback(mockContext);
    });
  });

  describe('渲染', () => {
    it('应该正确渲染组件', async () => {
      render(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByText('幻灯片移动工具')).toBeInTheDocument();
      });

      expect(screen.getByText('方法1: 移动指定幻灯片')).toBeInTheDocument();
      expect(screen.getByText('方法2: 移动当前选中的幻灯片')).toBeInTheDocument();
      expect(screen.getByText('方法3: 交换两张幻灯片位置')).toBeInTheDocument();
    });

    it('应该显示幻灯片信息概览', async () => {
      render(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByText(/总幻灯片数:/)).toBeInTheDocument();
      });

      expect(screen.getByText(/5/)).toBeInTheDocument();
    });

    it('应该显示幻灯片列表', async () => {
      render(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByText('Slide 1')).toBeInTheDocument();
      });

      expect(screen.getByText('Slide 2')).toBeInTheDocument();
      expect(screen.getByText('Slide 3')).toBeInTheDocument();
      expect(screen.getByText('Slide 4')).toBeInTheDocument();
      expect(screen.getByText('Slide 5')).toBeInTheDocument();
    });
  });

  describe('方法1: 移动指定幻灯片', () => {
    it('应该成功移动幻灯片', async () => {
      vi.mocked(pptTools.moveSlide).mockResolvedValue({
        success: true,
        message: '成功将幻灯片从位置 1 移动到位置 3',
        fromIndex: 1,
        toIndex: 3,
        totalSlides: 5,
      });

      render(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByLabelText('源位置:')).toBeInTheDocument();
      });

      const fromInput = screen.getByLabelText('源位置:') as HTMLInputElement;
      const toInput = screen.getByLabelText('目标位置:') as HTMLInputElement;
      const moveButton = screen.getByRole('button', { name: '移动幻灯片' });

      fireEvent.change(fromInput, { target: { value: '1' } });
      fireEvent.change(toInput, { target: { value: '3' } });
      fireEvent.click(moveButton);

      await waitFor(() => {
        expect(pptTools.moveSlide).toHaveBeenCalledWith({ fromIndex: 1, toIndex: 3 });
      });

      await waitFor(() => {
        expect(screen.getByText(/移动成功/)).toBeInTheDocument();
      });
    });

    it('应该显示错误信息当输入无效', async () => {
      render(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByLabelText('源位置:')).toBeInTheDocument();
      });

      const moveButton = screen.getByRole('button', { name: '移动幻灯片' });
      fireEvent.click(moveButton);

      await waitFor(() => {
        expect(screen.getByText(/请输入有效的源位置/)).toBeInTheDocument();
      });
    });

    it('应该禁用按钮当输入为空', async () => {
      render(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByRole('button', { name: '移动幻灯片' })).toBeDisabled();
      });
    });
  });

  describe('方法2: 移动当前幻灯片', () => {
    it('应该成功移动当前幻灯片', async () => {
      vi.mocked(pptTools.moveCurrentSlide).mockResolvedValue({
        success: true,
        message: '成功将当前幻灯片从位置 1 移动到位置 4',
        fromIndex: 1,
        toIndex: 4,
        totalSlides: 5,
      });

      render(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByLabelText('移动到位置:')).toBeInTheDocument();
      });

      const toInput = screen.getByLabelText('移动到位置:') as HTMLInputElement;
      const moveButton = screen.getByRole('button', { name: '移动当前幻灯片' });

      fireEvent.change(toInput, { target: { value: '4' } });
      fireEvent.click(moveButton);

      await waitFor(() => {
        expect(pptTools.moveCurrentSlide).toHaveBeenCalledWith(4);
      });

      await waitFor(() => {
        expect(screen.getByText(/移动成功/)).toBeInTheDocument();
      });
    });

    it('应该显示快速移动按钮', async () => {
      render(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByText('快速移动到:')).toBeInTheDocument();
      });

      expect(screen.getByRole('button', { name: '开头' })).toBeInTheDocument();
      expect(screen.getByRole('button', { name: '末尾' })).toBeInTheDocument();
    });

    it('应该通过快速按钮设置目标位置', async () => {
      render(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByRole('button', { name: '末尾' })).toBeInTheDocument();
      });

      const endButton = screen.getByRole('button', { name: '末尾' });
      fireEvent.click(endButton);

      const toInput = screen.getByLabelText('移动到位置:') as HTMLInputElement;
      expect(toInput.value).toBe('5');
    });
  });

  describe('方法3: 交换幻灯片', () => {
    it('应该成功交换两张幻灯片', async () => {
      vi.mocked(pptTools.swapSlides).mockResolvedValue({
        success: true,
        message: '成功交换位置 2 和位置 4 的幻灯片',
        fromIndex: 2,
        toIndex: 4,
        totalSlides: 5,
      });

      render(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByLabelText('第一张位置:')).toBeInTheDocument();
      });

      const index1Input = screen.getByLabelText('第一张位置:') as HTMLInputElement;
      const index2Input = screen.getByLabelText('第二张位置:') as HTMLInputElement;
      const swapButton = screen.getByRole('button', { name: '交换幻灯片' });

      fireEvent.change(index1Input, { target: { value: '2' } });
      fireEvent.change(index2Input, { target: { value: '4' } });
      fireEvent.click(swapButton);

      await waitFor(() => {
        expect(pptTools.swapSlides).toHaveBeenCalledWith(2, 4);
      });

      await waitFor(() => {
        expect(screen.getByText(/交换成功/)).toBeInTheDocument();
      });
    });

    it('应该显示错误信息当交换失败', async () => {
      vi.mocked(pptTools.swapSlides).mockResolvedValue({
        success: false,
        message: '两张幻灯片索引相同，无需交换',
        fromIndex: 2,
        toIndex: 2,
      });

      render(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByLabelText('第一张位置:')).toBeInTheDocument();
      });

      const index1Input = screen.getByLabelText('第一张位置:') as HTMLInputElement;
      const index2Input = screen.getByLabelText('第二张位置:') as HTMLInputElement;
      const swapButton = screen.getByRole('button', { name: '交换幻灯片' });

      fireEvent.change(index1Input, { target: { value: '2' } });
      fireEvent.change(index2Input, { target: { value: '2' } });
      fireEvent.click(swapButton);

      await waitFor(() => {
        expect(screen.getByText(/交换失败/)).toBeInTheDocument();
      });
    });
  });

  describe('刷新功能', () => {
    it('应该刷新幻灯片列表', async () => {
      render(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByRole('button', { name: /刷新列表/ })).toBeInTheDocument();
      });

      const refreshButton = screen.getByRole('button', { name: /刷新列表/ });
      fireEvent.click(refreshButton);

      await waitFor(() => {
        expect(pptTools.getAllSlidesInfo).toHaveBeenCalled();
      });
    });

    it('应该在操作成功后自动刷新', async () => {
      vi.mocked(pptTools.moveSlide).mockResolvedValue({
        success: true,
        message: '成功移动',
        fromIndex: 1,
        toIndex: 3,
        totalSlides: 5,
      });

      render(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByLabelText('源位置:')).toBeInTheDocument();
      });

      const fromInput = screen.getByLabelText('源位置:') as HTMLInputElement;
      const toInput = screen.getByLabelText('目标位置:') as HTMLInputElement;
      const moveButton = screen.getByRole('button', { name: '移动幻灯片' });

      fireEvent.change(fromInput, { target: { value: '1' } });
      fireEvent.change(toInput, { target: { value: '3' } });
      fireEvent.click(moveButton);

      await waitFor(() => {
        // 初始加载 + 操作后刷新 = 至少2次调用
        expect(pptTools.getAllSlidesInfo).toHaveBeenCalledTimes(2);
      });
    });
  });

  describe('错误处理', () => {
    it('应该处理 API 错误', async () => {
      vi.mocked(pptTools.moveSlide).mockRejectedValue(new Error('API Error'));

      render(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByLabelText('源位置:')).toBeInTheDocument();
      });

      const fromInput = screen.getByLabelText('源位置:') as HTMLInputElement;
      const toInput = screen.getByLabelText('目标位置:') as HTMLInputElement;
      const moveButton = screen.getByRole('button', { name: '移动幻灯片' });

      fireEvent.change(fromInput, { target: { value: '1' } });
      fireEvent.change(toInput, { target: { value: '3' } });
      fireEvent.click(moveButton);

      await waitFor(() => {
        expect(screen.getByText(/移动失败.*API Error/)).toBeInTheDocument();
      });
    });

    it('应该处理加载幻灯片信息失败', async () => {
      vi.mocked(pptTools.getAllSlidesInfo).mockRejectedValue(new Error('Load Error'));

      render(<SlideMove />);

      await waitFor(() => {
        expect(screen.getByText(/加载幻灯片信息失败/)).toBeInTheDocument();
      });
    });
  });
});
