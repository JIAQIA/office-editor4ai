---
trigger: glob
description: 
globs: src/**/*.tsx
---

因为Office AddIn界面在Office右侧，更细长，因此需要使用移动端的界面设计思路，而不是Web端。

同时主工作页面应该设置一个最小宽度。如果宽度过窄，可以横向滚动，但如果宽了，应该自适应布局

界面主要使用 fluentui 开发