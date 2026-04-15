% 1. 读取数据
data = readtable('机器学习.xlsx');


% 3. 分离预测目标 Y 和特征变量
predictorVars = data.Properties.VariableNames;
predictorVars(strcmp(predictorVars, 'Y')) = [];  % 移除 Y

% 4. 建立随机森林模型（回归）
rng(1);
nTrees = 1000;
rfModel = TreeBagger(nTrees, data, 'Y', ...
    'Method', 'regression', ...
    'OOBPrediction', 'On', ...
    'OOBPredictorImportance', 'On');

% 5. 提取特征重要性（解释率）
importance = rfModel.OOBPermutedPredictorDeltaError;
[sortedImp, idx] = sort(importance, 'descend');
sortedFeatures = predictorVars(idx);

% 显示变量对 Y 的解释率
disp('特征对 Y 的解释率（从高到低）：');
for i = 1:length(sortedImp)
    fprintf('%s: %.4f\n', sortedFeatures{i}, sortedImp(i));
end

% 可视化
figure;
bar(sortedImp);
set(gca, 'XTickLabel', sortedFeatures, 'XTickLabelRotation', 45);
ylabel('解释率');
title('随机森林特征重要性');

% 6. 用模型进行预测（以原数据为例）
predictedY = predict(rfModel, data);

% 显示部分预测结果
disp('前5个样本的预测值：');
disp(predictedY(1:5));
% 7. 计算 R?（决定系数）
trueY = data.Y;
SS_res = sum((trueY - predictedY).^2);                  % 残差平方和
SS_tot = sum((trueY - mean(trueY)).^2);                 % 总平方和
R_squared = 1 - SS_res / SS_tot;

fprintf('模型的 R? 值为：%.4f\n', R_squared);
% ...（前面的随机森林代码保持不变）...

% 在随机森林代码之后添加以下内容

% ...（前面的随机森林代码保持不变）...

%% 绘制Y与各个变量的散点图并拟合曲线
% 获取所有特征变量名（排除Y）
predictorVars = data.Properties.VariableNames;
predictorVars(strcmp(predictorVars, 'Y')) = [];

% 排除分类变量（如IGBP）
continuousVars = predictorVars;
for i = length(continuousVars):-1:1
    if iscategorical(data.(continuousVars{i}))
        continuousVars(i) = [];
    end
end

% 创建新图形窗口
figure('Units', 'normalized', 'Position', [0.1, 0.1, 0.8, 0.8]);

% 计算子图布局
numVars = length(continuousVars);
numCols = 4; % 每行4个子图
numRows = ceil(numVars / numCols);

for i = 1:numVars
    varName = continuousVars{i};
    x = data.(varName);
    y = data.Y;
    
    % 移除缺失值
    validIdx = ~isnan(x) & ~isnan(y);
    x_clean = x(validIdx);
    y_clean = y(validIdx);
    
    % 创建子图
    subplot(numRows, numCols, i);
    
    % 绘制散点图
    scatter(x_clean, y_clean, 15, 'filled', 'MarkerFaceAlpha', 0.6);
    hold on;
    
    % 尝试线性拟合
    linearModel = fitlm(x_clean, y_clean, 'linear');
    linearR2 = linearModel.Rsquared.Adjusted;
    pValue_linear = linearModel.coefTest;
    
    % 尝试二次拟合
    quadModel = fitlm([x_clean, x_clean.^2], y_clean, 'linear'); % 手动添加二次项
    quadR2 = quadModel.Rsquared.Adjusted;
    pValue_quad = quadModel.coefTest;
    
    % 选择最佳模型（二次拟合优度提高超过0.02）
    if (quadR2 - linearR2) > 0.02
        model = quadModel;
        modelType = 'quadratic';
        R2 = quadR2;
        pValue = pValue_quad;
        
        % 获取二次模型的预测和置信区间
        x_range = linspace(min(x_clean), max(x_clean), 100)';
        [y_pred, ci] = predict(model, [x_range, x_range.^2]);
    else
        model = linearModel;
        modelType = 'linear';
        R2 = linearR2;
        pValue = pValue_linear;
        
        % 获取线性模型的预测和置信区间
        x_range = linspace(min(x_clean), max(x_clean), 100)';
        [y_pred, ci] = predict(model, x_range);
    end
    
    % 绘制置信带
    fill([x_range; flipud(x_range)], [ci(:,1); flipud(ci(:,2))], ...
         [0.7 0.7 1], 'EdgeColor', 'none', 'FaceAlpha', 0.3);
    
    % 绘制拟合曲线
    plot(x_range, y_pred, 'r-', 'LineWidth', 2);
    
    % 标注R?和P值
    if pValue < 0.001
        pText = 'P < 0.001';
    else
        pText = sprintf('P = %.3f', pValue);
    end
    
    text(0.05, 0.95, ...
        sprintf('Adj R? = %.3f\n%s', R2, pText), ...
        'Units', 'normalized', ...
        'VerticalAlignment', 'top', ...
        'BackgroundColor', 'white');
    
    % 添加标题和标签
    title(sprintf('%s (%s)', varName, modelType));
    xlabel(varName);
    ylabel('Y');
    grid on;
    hold off;
end

%% 绘制预测Y与观测Y的散点图（兼容旧版本）
figure;
scatter(trueY, predictedY, 40, 'filled', 'MarkerFaceAlpha', 0.6);
hold on;

% 添加1:1参考线
maxVal = max([trueY; predictedY]);
minVal = min([trueY; predictedY]);
plot([minVal maxVal], [minVal maxVal], 'k--', 'LineWidth', 1.5);

% 计算拟合线和统计指标
mdl = fitlm(trueY, predictedY);
R2 = mdl.Rsquared.Ordinary;
pValue = mdl.coefTest;

% 添加回归线
x_range = linspace(minVal, maxVal, 100);
y_pred = predict(mdl, x_range');
plot(x_range, y_pred, 'r-', 'LineWidth', 2);

% 标注统计信息
if pValue < 0.001
    pText = 'P < 0.001';
else
    pText = sprintf('P = %.3f', pValue);
end

text(0.05, 0.95, sprintf('R? = %.3f\n%s', R2, pText), ...
    'Units', 'normalized', 'FontSize', 12, ...
    'VerticalAlignment', 'top', 'BackgroundColor', 'white');

% 添加标签和标题
xlabel('观测 Y');
ylabel('预测 Y');
title('随机森林模型预测性能');
grid on;
axis equal;
box on;

% 添加图例
legend('数据点', '1:1 参考线', '回归线', 'Location', 'best');
hold off;

%% 读取全球预测数据
predData = readtable('C:\Users\zhangqin\Desktop\博士后\SM threshold\土壤tiff文件\prediction.csv');

% 提取经纬度
lon = predData{:,1};  % 第一列是经度
lat = predData{:,2};  % 第二列是纬度

% 提取预测所需的变量列（与训练数据匹配的列名）
X_pred = predData(:, predictorVars);  % 保证变量名顺序一致

% 处理分类变量（如果有）：转换为categorical（与训练模型一致）
for i = 1:numel(predictorVars)
    if iscategorical(data.(predictorVars{i})) && ~iscategorical(X_pred.(predictorVars{i}))
        X_pred.(predictorVars{i}) = categorical(X_pred.(predictorVars{i}));
    end
end

% 使用训练好的模型进行预测
Y_pred_map = predict(rfModel, X_pred);

%% 可视化：绘制全球预测地图（散点图方式）
figure('Name', '全球预测Y值分布图');
scatter(lon, lat, 10, Y_pred_map, 'filled');
set(gca, 'YDir', 'normal');
xlabel('经度');
ylabel('纬度');
title('随机森林模型预测的全球Y值');
colorbar;
colormap(parula);  % 可选 colormap，如 parula, jet, viridis 等
axis equal;
grid on;

% 如果需要限制坐标范围（例如仅绘制陆地或特定区域），可添加：
% xlim([-180 180]);
% ylim([-60 90]);
