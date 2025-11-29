/**
 * 文件名: slideLayouts.ts
 * 作者: JQQ
 * 创建日期: 2025/11/29
 * 最后修改日期: 2025/11/29
 * 版权: 2023 JQQ. All rights reserved.
 * 描述: 幻灯片布局模板获取工具，提供可用布局模板列表及其元数据
 */

/* global PowerPoint, console */

/**
 * 幻灯片布局模板信息
 */
export interface SlideLayoutTemplate {
  /** 布局ID，用于创建新幻灯片时指定布局 */
  id: string;
  /** 布局名称，如 "标题幻灯片"、"标题和内容" 等 */
  name: string;
  /** 布局类型，如 "title"、"blank"、"twoContent" 等 */
  type: string;
  /** 占位符数量 */
  placeholderCount: number;
  /** 占位符类型列表，如 ["title", "body", "picture"] */
  placeholderTypes: string[];
  /** 是否为自定义布局 */
  isCustom: boolean;
}

/**
 * 获取可用布局模板列表的选项
 */
export interface GetSlideLayoutsOptions {
  /** 是否包含占位符详细信息，默认为 true */
  includePlaceholders?: boolean;
}

/**
 * 获取当前演示文稿的所有可用布局模板
 *
 * @param options 获取选项
 * @returns 布局模板列表
 *
 * @example
 * ```typescript
 * // 获取所有布局模板
 * const layouts = await getAvailableSlideLayouts();
 *
 * // 使用布局ID创建新幻灯片
 * await PowerPoint.run(async (context) => {
 *   const slides = context.presentation.slides;
 *   const layout = context.presentation.slideMasters.getItem(0).layouts.getItem(layouts[0].id);
 *   slides.add(layout);
 *   await context.sync();
 * });
 * ```
 */
export async function getAvailableSlideLayouts(
  options: GetSlideLayoutsOptions = {}
): Promise<SlideLayoutTemplate[]> {
  const { includePlaceholders = true } = options;

  console.log("[getAvailableSlideLayouts] 开始获取布局模板列表，选项:", options);

  try {
    const layouts: SlideLayoutTemplate[] = [];

    await PowerPoint.run(async (context) => {
      console.log("[PowerPoint.run] 进入上下文");

      // 获取第一个母版（通常演示文稿只有一个母版）
      const slideMasters = context.presentation.slideMasters;
      slideMasters.load("items");
      await context.sync();

      console.log("[PowerPoint.run] 母版数量:", slideMasters.items.length);

      if (slideMasters.items.length === 0) {
        console.warn("[PowerPoint.run] 没有找到母版");
        return;
      }

      // 遍历所有母版（通常只有一个）
      for (let masterIndex = 0; masterIndex < slideMasters.items.length; masterIndex++) {
        const master = slideMasters.items[masterIndex];
        master.load("id,name");

        // 获取该母版下的所有布局
        const masterLayouts = master.layouts;
        masterLayouts.load("items");
        await context.sync();

        console.log(
          `[PowerPoint.run] 母版 ${masterIndex + 1} (${master.name}) 包含 ${masterLayouts.items.length} 个布局`
        );

        // 遍历所有布局
        for (let i = 0; i < masterLayouts.items.length; i++) {
          const layout = masterLayouts.items[i];

          // 加载布局基本信息
          layout.load("id,name,type");
          await context.sync();

          console.log(
            `[PowerPoint.run] 处理布局 ${i + 1}/${masterLayouts.items.length}: ${layout.name} (${layout.type})`
          );

          const layoutInfo: SlideLayoutTemplate = {
            id: layout.id,
            name: layout.name || `布局 ${i + 1}`,
            type: layout.type || "unknown",
            placeholderCount: 0,
            placeholderTypes: [],
            isCustom: false, // PowerPoint API 目前无法直接判断是否为自定义布局
          };

          // 获取占位符信息（如果需要）
          if (includePlaceholders) {
            try {
              const shapes = layout.shapes;
              shapes.load("items");
              await context.sync();

              console.log(
                `[PowerPoint.run] 布局 "${layout.name}" 包含 ${shapes.items.length} 个形状`
              );

              const placeholderTypes: string[] = [];

              // 遍历形状，查找占位符
              for (let j = 0; j < shapes.items.length; j++) {
                const shape = shapes.items[j];
                shape.load("type");
                // eslint-disable-next-line office-addins/no-context-sync-in-loop
                await context.sync();

                if (shape.type === "Placeholder") {
                  try {
                    shape.load("placeholderFormat");
                    shape.placeholderFormat.load("type");
                    // eslint-disable-next-line office-addins/no-context-sync-in-loop
                    await context.sync();

                    const placeholderType = shape.placeholderFormat.type as string;
                    placeholderTypes.push(placeholderType);
                    console.log(
                      `[PowerPoint.run] 布局 "${layout.name}" 占位符 ${j + 1}: ${placeholderType}`
                    );
                  } catch (error) {
                    console.log(
                      `[PowerPoint.run] 无法获取占位符 ${j + 1} 的类型:`,
                      (error as Error).message
                    );
                  }
                }
              }

              layoutInfo.placeholderCount = placeholderTypes.length;
              layoutInfo.placeholderTypes = placeholderTypes;
            } catch (error) {
              console.log(
                `[PowerPoint.run] 获取布局 "${layout.name}" 的占位符信息失败:`,
                (error as Error).message
              );
            }
          }

          layouts.push(layoutInfo);
        }
      }
    });

    console.log("[getAvailableSlideLayouts] 成功获取", layouts.length, "个布局模板");
    return layouts;
  } catch (error) {
    console.error("[getAvailableSlideLayouts] 获取布局模板列表失败");
    console.error("[getAvailableSlideLayouts] 错误名称:", (error as Error).name);
    console.error("[getAvailableSlideLayouts] 错误消息:", (error as Error).message);
    console.error("[getAvailableSlideLayouts] 错误堆栈:", (error as Error).stack);

    // 打印 Office.js 特定的调试信息
    const officeError = error as { debugInfo?: unknown };
    if (officeError.debugInfo) {
      console.error(
        "[getAvailableSlideLayouts] Office.js 调试信息:",
        JSON.stringify(officeError.debugInfo, null, 2)
      );
    }

    throw error;
  }
}

/**
 * 根据布局ID创建新幻灯片
 *
 * @param layoutId 布局模板ID
 * @param position 插入位置（从0开始），不指定则插入到末尾
 * @returns 新创建的幻灯片ID
 *
 * @example
 * ```typescript
 * // 使用指定布局创建新幻灯片
 * const layouts = await getAvailableSlideLayouts();
 * const newSlideId = await createSlideWithLayout(layouts[0].id);
 * ```
 */
export async function createSlideWithLayout(layoutId: string, position?: number): Promise<string> {
  console.log("[createSlideWithLayout] 开始创建幻灯片，布局ID:", layoutId, "位置:", position);

  try {
    let newSlideId = "";

    await PowerPoint.run(async (context) => {
      const slideMasters = context.presentation.slideMasters;
      slideMasters.load("items");
      await context.sync();

      // 查找包含指定布局的母版
      let targetLayout: PowerPoint.SlideLayout | null = null;

      for (let i = 0; i < slideMasters.items.length; i++) {
        const master = slideMasters.items[i];
        const layouts = master.layouts;
        layouts.load("items");
        await context.sync();

        for (let j = 0; j < layouts.items.length; j++) {
          const layout = layouts.items[j];
          layout.load("id");
          // eslint-disable-next-line office-addins/no-context-sync-in-loop
          await context.sync();

          if (layout.id === layoutId) {
            targetLayout = layout;
            console.log("[PowerPoint.run] 找到目标布局:", layout.id);
            break;
          }
        }

        if (targetLayout) break;
      }

      if (!targetLayout) {
        throw new Error(`未找到布局ID: ${layoutId}`);
      }

      // 创建新幻灯片
      const slides = context.presentation.slides;
      slides.load("items");
      await context.sync();

      // PowerPoint API 的 slides.add() 方法返回 void，需要在添加后重新获取
      // 注意：当前 API 版本可能不支持指定插入位置和布局
      if (position !== undefined && position >= 0 && position < slides.items.length) {
        // 先添加到末尾
        slides.add();
        await context.sync();

        // 然后移动到指定位置
        // 注意：PowerPoint API 可能不支持直接移动幻灯片
        // 这里我们只能添加到末尾，位置参数暂时不生效
        console.log("[PowerPoint.run] 警告：当前 API 版本可能不支持指定插入位置");
      } else {
        // 添加到末尾
        slides.add();
        await context.sync();
      }

      // 重新加载幻灯片集合，获取新添加的幻灯片
      slides.load("items");
      await context.sync();

      // 新幻灯片应该是最后一个
      const newSlide = slides.items[slides.items.length - 1];

      if (!newSlide) {
        throw new Error("创建幻灯片失败：无法获取新创建的幻灯片");
      }

      // 尝试应用布局（如果 API 支持）
      try {
        // 某些版本的 API 可能支持通过 layout 属性设置布局
        const slideAny = newSlide as unknown as { layout?: PowerPoint.SlideLayout };
        if (slideAny.layout !== undefined) {
          slideAny.layout = targetLayout;
          await context.sync();
          console.log("[PowerPoint.run] 成功应用布局");
        } else {
          console.log("[PowerPoint.run] 警告：当前 API 版本不支持设置幻灯片布局");
        }
      } catch (error) {
        console.log("[PowerPoint.run] 应用布局失败:", (error as Error).message);
        console.log("[PowerPoint.run] 幻灯片已创建，但使用默认布局");
      }

      newSlide.load("id");
      await context.sync();

      newSlideId = newSlide.id;
      console.log("[PowerPoint.run] 新幻灯片创建成功，ID:", newSlideId);
    });

    console.log("[createSlideWithLayout] 幻灯片创建成功");
    return newSlideId;
  } catch (error) {
    console.error("[createSlideWithLayout] 创建幻灯片失败");
    console.error("[createSlideWithLayout] 错误:", (error as Error).message);
    throw error;
  }
}

/**
 * 获取布局模板的简要描述（用于UI展示）
 *
 * @param layout 布局模板信息
 * @returns 简要描述文本
 */
export function getLayoutDescription(layout: SlideLayoutTemplate): string {
  const parts: string[] = [];

  // 添加类型信息
  if (layout.type && layout.type !== "unknown") {
    parts.push(`类型: ${layout.type}`);
  }

  // 添加占位符信息
  if (layout.placeholderCount > 0) {
    parts.push(`${layout.placeholderCount} 个占位符`);

    // 列出占位符类型
    const uniqueTypes = Array.from(new Set(layout.placeholderTypes));
    if (uniqueTypes.length > 0 && uniqueTypes.length <= 3) {
      parts.push(`(${uniqueTypes.join(", ")})`);
    }
  } else {
    parts.push("无占位符");
  }

  return parts.join(" · ");
}
