% 批量处理文件夹中的多个 .csv 文件
folder_path = 'C:\Users\zhangqin\Desktop\daily\Hour\all\10个代表'; % 替换为您的文件夹路径
file_list = dir(fullfile(folder_path, '*.csv')); % 获取所有 .csv 文件

% 创建 Word 对象
wordApp = actxserver('Word.Application');
wordApp.Visible = true;  % 设置为 true 可显示 Word

% 创建一个新的文档
doc = wordApp.Documents.Add;

% 初始化结果存储
results = cell(length(file_list), 4);

% 遍历文件夹中的每个文件
for file_idx = 5:5
    % 读取当前文件
    file_name = file_list(file_idx).name;
    file_path = fullfile(folder_path, file_name);
    data = readtable(file_path);

 % 提取需要的列
    temperature = data.TA_F;
    ER = data.NEE_VUT_REF;
    swc = data.SWC_F_MDS_1; % 选择合适的 SWC 列
    night_flag = data.NIGHT; % 夜间标志
    precipitation = data.P_F; % 降水数据列
    ERQ = data.NEE_VUT_REF_QC; % ER 质量控制

    % 定义 validIdx，最初设定为全部数据有效
    validIdx = true(size(temperature));

    % 查找降水事件：通过连续不为0的降水记录确定降水事件
    exclude_idx = false(size(precipitation));
    i = 1;

    while i <= length(precipitation)
        if precipitation(i) > 0
            % 降水事件开始
            start_idx = i;
            cumulative_precipitation = 0; % 初始化累计降水量

            % 向后查找，直到降水结束
            while i <= length(precipitation) && precipitation(i) > 0
                cumulative_precipitation = cumulative_precipitation + precipitation(i);
                i = i + 1;
            end
            end_idx = i - 1; % 降水事件的结束索引

            % 判断降水事件的累计降水量是否超过 15mm
            if cumulative_precipitation >= 15
                % 排除降水事件及其之后 48 小时的数据（96 个半小时）
                exclude_idx(start_idx:end_idx) = true;
                if end_idx + 144 <= length(precipitation)
                    exclude_idx(end_idx + 1:end_idx + 144) = true; % 排除 48 小时内的数据
                else
                    exclude_idx(end_idx + 1:end) = true; % 如果数据不足 48 小时，则排除剩余所有数据
                end
            end
        else
            i = i + 1;
        end
    end

    % 更新 validIdx，排除降水相关数据
    validIdx = validIdx & ~exclude_idx;

    % 去除无效值并仅保留夜间数据
    validIdx = validIdx & temperature ~= -9999 & ER >= 0 & swc ~= -9999 & night_flag == 1 & ERQ ~= 3 & ERQ ~= 2;

    % 使用 validIdx 对数据进行筛选
    temperature = temperature(validIdx);
    ER = ER(validIdx);
    swc = swc(validIdx);

    % 计算整体温度范围长度
    overall_temp_range = max(temperature) - min(temperature);
    min_temp_range_threshold = 0.5 * overall_temp_range; % 50% 阈值

    % 定义 SWC 间隔（以 5 为长度，1 为间隔）
    swc_min = floor(min(swc));
    swc_max = ceil(max(swc));
    swc_start_points = swc_min:1:(swc_max - 5); % 滑动窗口的起点
    swc_midpoints = swc_start_points + 2.5; % 每个区间的中点

    % 用于存储当前文件的 SWC 拟合参数
    swc_fit_results = [];
    exp_10alpha_values = []; % 存储 e^(10*α) 的值

    % 按 SWC 分组分析
    for swc_idx = 1:length(swc_start_points)
        % 当前 SWC 区间
        swc_start = swc_start_points(swc_idx);
        swc_end = swc_start + 5;
        swc_mask = swc >= swc_start & swc < swc_end;

        if sum(swc_mask) > 0
            % 获取当前 SWC 区间内的数据
            temp_subset = temperature(swc_mask);
            ER_subset = ER(swc_mask);

            % 计算当前 SWC 区间的温度范围
            temp_min = min(temp_subset);
            temp_max = max(temp_subset);
            temp_range = temp_max - temp_min;

            % 如果温度范围太小，跳过该区间
            if temp_range < min_temp_range_threshold
                exp_10alpha_values = [exp_10alpha_values; NaN];
                continue; % 跳过这个 SWC 区间
            end

            % 拟合 ER = γe^(αT)
            fit_func = @(params, T) params(1) * exp(params(2) * T);
            initial_params = [1, 0.1];
            error_func = @(params) sum((fit_func(params, temp_subset) - ER_subset).^2);
            optimal_params = fminsearch(error_func, initial_params);

            alpha = optimal_params(2);
            exp_10alpha = exp(10 * alpha);
            exp_10alpha_values = [exp_10alpha_values; exp_10alpha];
        else
            exp_10alpha_values = [exp_10alpha_values; NaN];
        end
    end

    % 计算 SWC 平均值
    swc_mean = mean(swc);

    % 找到 e^(10*α) 最大值对应的 SWC 中间值
    [max_exp_10alpha, max_idx] = max(exp_10alpha_values);
    if ~isempty(max_idx) && ~isnan(max_exp_10alpha)
        swc_max_exp_10alpha = swc_midpoints(max_idx);
    else
        swc_max_exp_10alpha = NaN;
    end

    % 计算该区间的 SWC 中位数
    if ~isnan(swc_max_exp_10alpha)
        swc_mask = swc >= swc_start_points(max_idx) & swc < (swc_start_points(max_idx) + 5);
        if sum(swc_mask) > 0
            swc_median = median(swc(swc_mask));
        else
            swc_median = NaN;
        end
    else
        swc_median = NaN;
    end

    % 计算百分位数
    if ~isnan(swc_max_exp_10alpha)
        prctile_rank = sum(swc < swc_max_exp_10alpha) / length(swc) * 100;
    else
        prctile_rank = NaN;
    end

    % 提取不带扩展名的文件名
    [~, baseFileName, ~] = fileparts(file_name);

    % 存储结果
    results{file_idx, 1} = baseFileName;
    results{file_idx, 2} = swc_mean;
    results{file_idx, 3} = swc_max_exp_10alpha;
    results{file_idx, 4} = prctile_rank;
end

% 在Word中创建表格
tableRange = doc.Range(doc.Content.End-1, doc.Content.End);
numRows = length(file_list) + 1; % 包含标题行
numCols = 4;
table = doc.Tables.Add(tableRange, numRows, numCols);
table.Borders.Enable = 1; % 启用表格边框

% 设置标题行
table.Cell(1, 1).Range.Text = '站点名';
table.Cell(1, 2).Range.Text = 'SWCmean';
table.Cell(1, 3).Range.Text = 'swc_max_exp_10alpha';
table.Cell(1, 4).Range.Text = 'prctile_rank (%)';

% 填充数据行
for row = 1:length(file_list)
    for col = 1:4
        currentData = results{row, col};
        if isnumeric(currentData)
            if isnan(currentData)
                textValue = 'N/A';
            else
                switch col
                    case {2, 3}
                        textValue = sprintf('%.4f', currentData);
                    case 4
                        textValue = sprintf('%.2f%%', currentData);
                    otherwise
                        textValue = num2str(currentData);
                end
            end
        else
            textValue = currentData;
        end
        table.Cell(row+1, col).Range.Text = textValue;
    end
end

% 调整表格格式（可选）
table.AutoFitBehavior(1); % 自动调整列宽
table.Rows.Item(1).Shading.BackgroundPatternColor = 15773696; % 设置标题行背景色

% 保存文档（可选）
% doc.SaveAs2(fullfile(folder_path, 'Results.docx'));