/* global PowerPoint console */

/**
 * 插入文本框到幻灯片
 * Insert text box to slide
 * @param text 文本内容 / Text content
 * @param left X坐标（可选，默认居中），单位：磅 / X coordinate (optional, centered by default), unit: points
 * @param top Y坐标（可选，默认居中），单位：磅 / Y coordinate (optional, centered by default), unit: points
 */
export async function insertText(text: string, left?: number, top?: number) {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      
      // 如果指定了位置参数，使用指定位置；否则使用默认位置
      // If position parameters are specified, use them; otherwise use default position
      let textBox;
      if (left !== undefined && top !== undefined) {
        // PowerPoint API: addTextBox(text, left, top, width, height)
        // 默认宽度300磅，高度100磅 / Default width 300 points, height 100 points
        textBox = slide.shapes.addTextBox(text, {
          left: left,
          top: top,
          width: 300,
          height: 100
        });
      } else {
        textBox = slide.shapes.addTextBox(text);
      }
      
      textBox.fill.setSolidColor("white");
      textBox.lineFormat.color = "black";
      textBox.lineFormat.weight = 1;
      textBox.lineFormat.dashStyle = PowerPoint.ShapeLineDashStyle.solid;
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}
