import React, { useState, useEffect } from 'react';
import { Play, AlertCircle, CheckCircle, RotateCcw } from 'lucide-react';

const ExecutionComponent = () => {
  // 状态管理：idle(空闲), running(运行中), success(成功), error(失败)
  const [status, setStatus] = useState('idle');
  const [progress, setProgress] = useState(0);
  const [errorMsg, setErrorMsg] = useState('');
  
  // 模拟的数据输入 (用于在线编辑)
  const [inputData, setInputData] = useState(JSON.stringify({
    "task_name": "测试任务",
    "target": "" // 故意留空用于演示报错
  }, null, 4));

  // 运行/校验逻辑
  const handleRun = () => {
    // 1. 重置状态
    setStatus('running');
    setProgress(0);
    setErrorMsg('');

    // 模拟加载过程
    let currentProgress = 0;
    const interval = setInterval(() => {
      currentProgress += 10;
      setProgress(currentProgress);

      if (currentProgress >= 100) {
        clearInterval(interval);
        validateData(); // 加载完后进行校验
      }
    }, 50); // 速度快一点
  };

  // 校验逻辑
  const validateData = () => {
    try {
      const parsed = JSON.parse(inputData);
      
      // 模拟校验规则：必须有 target 字段且不为空
      if (!parsed.target) {
        throw new Error("校验失败: 'target' 字段不能为空。请在下方编辑器中修正。");
      }

      // 校验通过
      setStatus('success');
    } catch (err) {
      // 校验不通过
      setStatus('error');
      setErrorMsg(err.message || "JSON 格式错误");
      setProgress(0); // 进度条归零或保持黑色背景
    }
  };

  // 重置
  const handleReset = () => {
    setStatus('idle');
    setProgress(0);
    setErrorMsg('');
  };

  return (
    <div className="w-full max-w-2xl mx-auto p-6 space-y-6">
      
      {/* 标题 */}
      <div className="text-gray-700 font-bold text-lg">执行控制台</div>

      {/* --- 核心进度条区域 --- */}
      <div className="relative w-full h-12 rounded-lg overflow-hidden shadow-lg border border-gray-800">
        
        {/* 1. 背景层 (默认黑色) */}
        <div className="absolute inset-0 bg-black z-0"></div>

        {/* 2. 进度/成功层 (绿色) 
            逻辑：只有在 running 或 success 时显示，宽度动态变化
        */}
        <div 
          className="absolute inset-0 bg-green-500 z-10 transition-all duration-300 ease-out"
          style={{ width: `${status === 'error' ? 0 : progress}%` }}
        ></div>

        {/* 3. 内容交互层 (按钮和文字) - 必须置于最顶层 (z-20) */}
        <div className="absolute inset-0 z-20 flex items-center px-4 justify-between">
          
          {/* 左侧：运行按钮 / 状态图标 */}
          <button 
            onClick={status === 'idle' || status === 'error' ? handleRun : undefined}
            disabled={status === 'running' || status === 'success'}
            className="flex items-center gap-2 focus:outline-none group"
          >
            {/* 图标逻辑切换 */}
            {status === 'running' && (
               <span className="animate-spin text-white">⏳</span>
            )}

            {status === 'idle' && (
              <>
                <div className="w-8 h-8 rounded-full bg-white/20 flex items-center justify-center group-hover:bg-white/40 transition">
                  <Play className="w-4 h-4 text-white fill-current" />
                </div>
                <span className="text-white font-mono text-sm">点击运行</span>
              </>
            )}

            {status === 'success' && (
              <>
                <CheckCircle className="w-6 h-6 text-white" />
                <span className="text-white font-bold">执行成功</span>
              </>
            )}

            {/* 失败状态：显示感叹号 */}
            {status === 'error' && (
              <>
                <div className="w-8 h-8 rounded-full bg-red-600 flex items-center justify-center animate-pulse">
                  <AlertCircle className="w-5 h-5 text-white" />
                </div>
                <span className="text-red-500 font-bold ml-2">校验未通过</span>
              </>
            )}
          </button>

          {/* 右侧：重置按钮 (仅在结束状态显示) */}
          {(status === 'success' || status === 'error') && (
            <button 
              onClick={handleReset}
              className="text-white/70 hover:text-white flex items-center gap-1 text-sm"
            >
              <RotateCcw className="w-4 h-4" /> 重置
            </button>
          )}
        </div>
      </div>

      {/* --- 错误提示区域 --- */}
      {status === 'error' && (
        <div className="p-4 bg-red-50 border-l-4 border-red-500 text-red-700 rounded animate-fade-in-down">
          <p className="font-bold">Execution Failed</p>
          <p className="text-sm">{errorMsg}</p>
        </div>
      )}

      {/* --- 在线数据编辑 & 可视化定位 --- */}
      <div className="space-y-2">
        <div className="flex justify-between items-center">
          <label className="text-sm font-semibold text-gray-600">数据配置 (JSON)</label>
          <span className="text-xs text-gray-400">支持在线编辑实时生效</span>
        </div>
        
        <div className="relative">
          <textarea
            value={inputData}
            onChange={(e) => {
              setInputData(e.target.value);
              if (status === 'error') setStatus('idle'); // 编辑时重置错误状态
            }}
            className={`w-full h-48 p-4 font-mono text-sm bg-gray-50 rounded-lg border-2 focus:outline-none transition-colors resize-none
              ${status === 'error' 
                ? 'border-red-500 bg-red-50/10'  // 错误时：边框变红
                : 'border-gray-200 focus:border-blue-500' // 正常时：灰色/蓝色
              }`}
            spellCheck="false"
          />
          
          {/* 可视化定位提示 (简单的模拟) */}
          {status === 'error' && inputData.includes('"target": ""') && (
            <div className="absolute top-[4.5rem] right-4 text-xs text-red-500 bg-white px-2 py-1 rounded shadow border border-red-200">
              👈 错误定位: 值不能为空
            </div>
          )}
        </div>
      </div>

    </div>
  );
};

export default ExecutionComponent;
