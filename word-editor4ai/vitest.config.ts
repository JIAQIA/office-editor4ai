/**
 * 文件名: vitest.config.ts
 * 作者: JQQ
 * 创建日期: 2025/11/30
 * 最后修改日期: 2025/11/30
 * 版权: 2023 JQQ. All rights reserved.
 * 依赖: vitest
 * 描述: Vitest 测试配置文件 | Vitest test configuration file
 */

import { defineConfig } from 'vitest/config';
import path from 'path';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  test: {
    // 测试环境：使用 jsdom 模拟浏览器环境 | Test environment: use jsdom to simulate browser environment
    environment: 'jsdom',
    
    // 全局测试设置 | Global test setup
    globals: true,
    
    // 设置文件：在每个测试文件前运行 | Setup files: run before each test file
    setupFiles: ['./tests/setup.ts'],
    
    // 覆盖率配置 | Coverage configuration
    coverage: {
      provider: 'v8',
      reporter: ['text', 'json', 'html', 'lcov'],
      exclude: [
        'node_modules/',
        'tests/',
        'dist/',
        '**/*.config.ts',
        '**/*.config.js',
        '**/commands.ts', // Office.js 命令文件 | Office.js command files
      ],
      // 覆盖率阈值 | Coverage thresholds
      thresholds: {
        lines: 60,
        functions: 60,
        branches: 60,
        statements: 60,
      },
    },
    
    // 包含的测试文件模式 | Test file patterns to include
    include: ['tests/**/*.{test,spec}.{ts,tsx}'],
    
    // 排除的文件模式 | File patterns to exclude
    exclude: [
      'node_modules',
      'dist',
      '.idea',
      '.git',
      '.cache',
    ],
    
    // 测试超时时间（毫秒）| Test timeout (milliseconds)
    testTimeout: 10000,
    
    // 钩子超时时间（毫秒）| Hook timeout (milliseconds)
    hookTimeout: 10000,
  },
  
  resolve: {
    alias: {
      // 路径别名配置 | Path alias configuration
      '@': path.resolve(__dirname, './src'),
      '@components': path.resolve(__dirname, './src/taskpane/components'),
      '@commands': path.resolve(__dirname, './src/commands'),
    },
  },
});
