from flask import Flask, request, render_template, send_file, redirect, url_for, jsonify
import os
import pandas as pd
from datetime import datetime

# 导入共享模块
import analyze_exam_data

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB limit

# 确保上传目录存在
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

# 辅助函数：将DataFrame转换为JSON格式
def dataframe_to_json(df):
    """将DataFrame转换为适合前端显示的JSON格式"""
    # 转换为字典格式
    data = []
    for idx, row in df.iterrows():
        row_data = {'班级': str(idx)}
        for col in df.columns:
            value = row[col]
            # 处理NaN值
            if pd.isna(value):
                row_data[col] = None
            # 确保所有值都能序列化
            elif isinstance(value, (int, float, str, bool)):
                row_data[col] = value
            else:
                row_data[col] = str(value)
        data.append(row_data)
    
    # 获取列名
    columns = ['班级'] + list(df.columns)
    
    return {'columns': columns, 'data': data}

# 路由：分析API（返回JSON格式结果）
@app.route('/api/analyze', methods=['POST'])
def analyze_api():
    """处理文件上传并返回JSON格式的分析结果"""
    # 检查是否有文件上传
    if 'file' not in request.files:
        return jsonify({'error': '请选择文件'}), 400
    
    file = request.files['file']
    
    # 检查文件是否为空
    if file.filename == '':
        return jsonify({'error': '请选择文件'}), 400
    
    # 检查文件类型
    if file and (file.filename.endswith('.xls') or file.filename.endswith('.xlsx')):
        try:
            # 使用共享模块读取文件
            data = analyze_exam_data.read_excel_file(file)
            
            # 计算班级平均分
            class_averages = analyze_exam_data.calculate_class_averages(data)
            
            # 分类数据
            chinese_cols = analyze_exam_data.classify_by_subject(data)
            
            # 生成信息学科分析
            chinese_analysis = analyze_exam_data.generate_chinese_analysis(data, chinese_cols)
            
            # 保存分析结果到临时文件
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            temp_file = os.path.join(app.config['UPLOAD_FOLDER'], f'analysis_{timestamp}.xlsx')
            analyze_exam_data.save_results_to_excel(class_averages, info_analysis, general_analysis, temp_file, data)
            
            # 转换为相对URL路径
            relative_temp_file = f'/uploads/analysis_{timestamp}.xlsx'
            
            # 转换为JSON格式
            result = {
                'class_averages': dataframe_to_json(class_averages),
                'info_analysis': dataframe_to_json(info_analysis),
                'general_analysis': dataframe_to_json(general_analysis),
                'success': True,
                'temp_file': relative_temp_file  # 保存相对URL路径，用于后续导出
            }
            
            return jsonify(result)
            
        except Exception as e:
            error_message = f'处理文件时出错: {str(e)}'
            return jsonify({'error': error_message}), 500
    else:
        return jsonify({'error': '请选择Excel文件（.xls或.xlsx格式）'}), 400

# 路由：导出结果API
@app.route('/api/export', methods=['POST'])
def export_api():
    """处理文件上传并返回Excel文件"""
    # 检查是否有文件上传
    if 'file' not in request.files:
        return jsonify({'error': '请选择文件'}), 400
    
    file = request.files['file']
    
    # 检查文件是否为空
    if file.filename == '':
        return jsonify({'error': '请选择文件'}), 400
    
    # 检查文件类型
    if file and (file.filename.endswith('.xls') or file.filename.endswith('.xlsx')):
        try:
            # 使用共享模块读取文件
            data = analyze_exam_data.read_excel_file(file)
            
            # 计算班级平均分
            class_averages = analyze_exam_data.calculate_class_averages(data)
            
            # 分类数据
            chinese_cols = analyze_exam_data.classify_by_subject(data)
            
            chinese_analysis = analyze_exam_data.generate_chinese_analysis(data, chinese_cols)
                        
            # 生成Excel文件（使用字节流）
            excel_file = analyze_exam_data.save_results_to_excel_bytes(class_averages, chinese_analysis,  data)
            
            # 生成文件名
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f'分析结果_{timestamp}.xlsx'
            
            # 返回文件下载
            return send_file(
                excel_file,
                as_attachment=True,
                download_name=filename,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            
        except Exception as e:
            error_message = f'处理文件时出错: {str(e)}'
            return jsonify({'error': error_message}), 500
    else:
        return jsonify({'error': '请选择Excel文件（.xls或.xlsx格式）'}), 400

# 路由：首页
@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

# 路由：提供上传文件的下载
@app.route('/uploads/<filename>')
def download_file(filename):
    from flask import send_from_directory
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)

import shutil
import time

# 清理过期的临时文件
def cleanup_temp_files():
    """清理超过24小时的临时分析文件"""
    upload_folder = app.config['UPLOAD_FOLDER']
    if not os.path.exists(upload_folder):
        return
    
    current_time = time.time()
    for filename in os.listdir(upload_folder):
        file_path = os.path.join(upload_folder, filename)
        if os.path.isfile(file_path):
            # 检查文件修改时间
            if current_time - os.path.getmtime(file_path) > 24 * 3600:  # 24小时
                try:
                    os.remove(file_path)
                    print(f"已清理过期文件: {filename}")
                except Exception as e:
                    print(f"清理文件时出错: {e}")

# 在应用启动时清理临时文件
cleanup_temp_files()

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
