# -*- coding: utf-8 -*-
"""
AIDA64报告解析核心模块
"""

import re
import os
import json
import pandas as pd
from datetime import datetime
from typing import List, Dict, Tuple, Optional


class AIDA64Parser:
    """AIDA64报告解析器"""
    
    def __init__(self):
        self.section_patterns = {
            '系统概述': r'--------\[ 系统概述 \]-+\n(.*?)\n\n--------',
            '计算机名称': r'--------\[ 计算机名称 \]-+\n(.*?)\n\n--------',
            'SPD': r'--------\[ SPD \]-+\n(.*?)\n\n--------',
            '显示器': r'--------\[ 显示器 \]-+\n(.*?)\n\n--------',
            '逻辑驱动器': r'--------\[ 逻辑驱动器 \]-+\n(.*?)\n\n--------',
            '物理驱动器': r'--------\[ 物理驱动器 \]-+\n(.*?)\n\n--------',
            'Windows 网络': r'--------\[ Windows 网络 \]-+\n(.*?)\n\n--------',
            '已安装程序': r'--------\[ 已安装程序 \]-+\n(.*?)(?:\n\n|\n\nThe names of)'
        }
        
        # 预定义需要提取的项目
        self.standard_items = {
            '系统概述': [
                '计算机类型', '操作系统', '计算机名称', '用户名称', '登录域',
                '处理器名称', '主板名称', '主板芯片组', '系统内存',
                '显示适配器', '3D 加速器', '显示器',
                '存储控制器1', '存储控制器2', '硬盘驱动器1', '硬盘驱动器2', 
                '硬盘 SMART 状态',
                '主 IP 地址', '主 MAC 地址', '网络适配器1', '网络适配器2', '网络适配器3'
            ],
            'DMI': [
                'DMI BIOS 厂商', 'DMI BIOS 版本', 'DMI 系统制造商', 'DMI 系统产品',
                'DMI 系统版本', 'DMI 系统序列号', 'DMI 系统 UUID', 'DMI 主板制造商',
                'DMI 主板产品', 'DMI 主板版本', 'DMI 主板序列号', 'DMI 主机制造商',
                'DMI 主机版本', 'DMI 主机序列号', 'DMI 主机识别标签', 'DMI 主机类型'
            ],
            'SPD': [
                'DIMM1: 模块名称', 'DIMM1: 序列号', 'DIMM1: 制造日期', 'DIMM1: 模块容量',
                'DIMM1: 模块类型', 'DIMM1: 存取类型', 'DIMM1: 存取速度', 'DIMM1: 模块位宽',
                'DIMM1: 模块电压', 'DIMM1: 错误检测方式', 'DIMM1: DRAM 制造商',
                'DIMM1: DRAM Stepping', 'DIMM1: SDRAM Die Count',
                'DIMM3: 模块名称', 'DIMM3: 序列号', 'DIMM3: 制造日期', 'DIMM3: 模块容量',
                'DIMM3: 模块类型', 'DIMM3: 存取类型', 'DIMM3: 存取速度', 'DIMM3: 模块位宽',
                'DIMM3: 模块电压', 'DIMM3: 错误检测方式', 'DIMM3: DRAM 制造商',
                'DIMM3: SDRAM Die Count'
            ]
        }
    
    def parse_file(self, file_path: str, selected_items: List[str] = None) -> List[Dict]:
        """
        解析AIDA64报告文件
        
        Args:
            file_path: 文件路径
            selected_items: 选中的项目列表
        
        Returns:
            解析结果列表，每个元素是 {'项目': ..., '值': ...}
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            all_data = []
            
            # 解析系统概述
            if selected_items is None or any(item in self.standard_items['系统概述'] for item in selected_items):
                system_data = self._parse_system_summary(content, selected_items)
                all_data.extend(system_data)
            
            # 解析DMI信息（在系统概述中）
            if selected_items is None or any(item.startswith('DMI') for item in selected_items):
                dmi_data = self._parse_dmi_info(content, selected_items)
                all_data.extend(dmi_data)
            
            # 解析SPD信息
            if selected_items is None or any(item.startswith('DIMM') for item in selected_items):
                spd_data = self._parse_spd_info(content, selected_items)
                all_data.extend(spd_data)
            
            # 解析磁盘分区
            if selected_items is None or any(item.startswith(('C:', 'D:', 'E:', '分区')):
                    self._parse_disk_info(content, selected_items)
                all_data.extend(disk_data)
            
            # 解析网络信息
            if selected_items is None or any(item.startswith(('网络适配器', 'IP地址', 'MAC地址')):
                    self._parse_network_info(content, selected_items)
                all_data.extend(network_data)
            
            # 解析已安装程序
            if selected_items is None or '已安装程序' in (selected_items or []):
                software_data = self._parse_installed_software(content)
                all_data.extend(software_data)
            
            return all_data
            
        except Exception as e:
            raise Exception(f"解析文件时出错: {str(e)}")
    
    def _parse_system_summary(self, content: str, selected_items: List[str] = None) -> List[Dict]:
        """解析系统概述部分"""
        data = []
        
        # 查找系统概述部分
        summary_match = re.search(self.section_patterns['系统概述'], content, re.DOTALL)
        if not summary_match:
            return data
        
        summary_content = summary_match.group(1)
        lines = summary_content.strip().split('\n')
        
        current_section = None
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # 检查是否是新的小节
            if line.endswith(':'):
                current_section = line[:-1]
                continue
            
            # 解析键值对
            if ':' in line and current_section:
                parts = line.split(':', 1)
                if len(parts) == 2:
                    key = f"{current_section}: {parts[0].strip()}" if current_section != '计算机' else parts[0].strip()
                    value = parts[1].strip()
                    
                    # 重命名部分键以匹配标准格式
                    if key == "计算机: 计算机类型":
                        key = "计算机类型"
                    elif key == "计算机: 操作系统":
                        key = "操作系统"
                    elif key == "计算机: 计算机名称":
                        key = "计算机名称"
                    elif key == "计算机: 用户名称":
                        key = "用户名称"
                    elif key == "计算机: 登录域":
                        key = "登录域"
                    
                    # 检查是否在选中的项目中
                    if selected_items is None or key in selected_items:
                        data.append({'项目': key, '值': value})
        
        return data
    
    def _parse_dmi_info(self, content: str, selected_items: List[str] = None) -> List[Dict]:
        """解析DMI信息"""
        data = []
        
        # 在系统概述中查找DMI部分
        dmi_pattern = r'    DMI:\s*\n(.*?)(?:\n\n|\n    \S)'
        dmi_match = re.search(dmi_pattern, content, re.DOTALL)
        
        if dmi_match:
            dmi_content = dmi_match.group(1)
            lines = dmi_content.strip().split('\n')
            
            for line in lines:
                line = line.strip()
                if ':' in line:
                    parts = line.split(':', 1)
                    if len(parts) == 2:
                        key = parts[0].strip()
                        value = parts[1].strip()
                        
                        if selected_items is None or key in selected_items:
                            data.append({'项目': key, '值': value})
        
        return data
    
    def _parse_spd_info(self, content: str, selected_items: List[str] = None) -> List[Dict]:
        """解析SPD信息"""
        data = []
        
        spd_match = re.search(self.section_patterns['SPD'], content, re.DOTALL)
        if not spd_match:
            return data
        
        spd_content = spd_match.group(1)
        
        # 解析DIMM1
        dimm1_pattern = r'\[ DIMM1: (.*?) \]\s*\n\n(.*?)(?:\n\n\[|\n\n--------)'
        dimm1_match = re.search(dimm1_pattern, spd_content, re.DOTALL)
        
        if dimm1_match:
            dimm1_name = dimm1_match.group(1)
            dimm1_info = dimm1_match.group(2)
            
            lines = dimm1_info.strip().split('\n')
            for line in lines:
                line = line.strip()
                if ':' in line:
                    parts = line.split(':', 1)
                    if len(parts) == 2:
                        key = f"DIMM1: {parts[0].strip()}"
                        value = parts[1].strip()
                        
                        if selected_items is None or key in selected_items:
                            data.append({'项目': key, '值': value})
        
        # 解析DIMM3
        dimm3_pattern = r'\[ DIMM3: (.*?) \]\s*\n\n(.*?)(?:\n\n|\Z)'
        dimm3_match = re.search(dimm3_pattern, spd_content, re.DOTALL)
        
        if dimm3_match:
            dimm3_info = dimm3_match.group(2)
            lines = dimm3_info.strip().split('\n')
            
            for line in lines:
                line = line.strip()
                if ':' in line:
                    parts = line.split(':', 1)
                    if len(parts) == 2:
                        key = f"DIMM3: {parts[0].strip()}"
                        value = parts[1].strip()
                        
                        if selected_items is None or key in selected_items:
                            data.append({'项目': key, '值': value})
        
        return data
    
    def _parse_disk_info(self, content: str, selected_items: List[str] = None) -> List[Dict]:
        """解析磁盘分区信息"""
        data = []
        
        # 查找逻辑驱动器部分
        logical_match = re.search(self.section_patterns['逻辑驱动器'], content, re.DOTALL)
        if logical_match:
            logical_content = logical_match.group(1)
            lines = logical_content.strip().split('\n')
            
            for line in lines:
                line = line.strip()
                if line and '本地驱动器' in line:
                    # 解析分区信息
                    parts = line.split()
                    if len(parts) >= 7:
                        drive = parts[0]
                        filesystem = parts[2]
                        total_size = parts[3]
                        used_space = parts[4]
                        free_space = parts[5]
                        usage = parts[6]
                        
                        data.append({'项目': f'{drive} 文件系统', '值': filesystem})
                        data.append({'项目': f'{drive} 总大小', '值': total_size})
                        data.append({'项目': f'{drive} 已用空间', '值': used_space})
                        data.append({'项目': f'{drive} 可用空间', '值': free_space})
                        data.append({'项目': f'{drive} 使用率', '值': usage})
        
        return data
    
    def _parse_network_info(self, content: str, selected_items: List[str] = None) -> List[Dict]:
        """解析网络信息"""
        data = []
        
        network_match = re.search(self.section_patterns['Windows 网络'], content, re.DOTALL)
        if not network_match:
            return data
        
        network_content = network_match.group(1)
        
        # 解析每个网络适配器
        adapters = re.split(r'\n\n  \[ ', network_content)
        
        for adapter in adapters:
            if not adapter.strip():
                continue
            
            lines = adapter.strip().split('\n')
            adapter_name = None
            
            for line in lines:
                line = line.strip()
                if '网络适配器:' in line:
                    adapter_name = line.split(':', 1)[1].strip()
                    data.append({'项目': '网络适配器', '值': adapter_name})
                elif ':' in line and adapter_name:
                    parts = line.split(':', 1)
                    if len(parts) == 2:
                        key = parts[0].strip()
                        value = parts[1].strip()
                        
                        # 只添加关键的网络信息
                        if key in ['IP 地址/子网掩码', '硬件地址(MAC)', '连接速度']:
                            full_key = f"{adapter_name}: {key}"
                            if selected_items is None or full_key in selected_items:
                                data.append({'项目': full_key, '值': value})
        
        return data
    
    def _parse_installed_software(self, content: str) -> List[Dict]:
        """解析已安装程序"""
        data = []
        
        software_match = re.search(self.section_patterns['已安装程序'], content, re.DOTALL)
        if not software_match:
            return data
        
        software_content = software_match.group(1)
        lines = software_content.strip().split('\n')
        
        for line in lines:
            line = line.strip()
            if line:
                # 尝试匹配软件名和版本
                # 格式: 软件名 版本 ...
                parts = line.split()
                if len(parts) >= 2:
                    # 尝试找到版本号开始的位置
                    for i in range(1, len(parts)):
                        if re.match(r'[\d\.]+', parts[i]):
                            software_name = ' '.join(parts[:i])
                            version = parts[i]
                            
                            # 清理版本号
                            version = re.sub(r'[^\d\.]', '', version)
                            
                            data.append({'项目': software_name, '值': version})
                            break
        
        return data
    
    def parse_multiple_files(self, file_paths: List[str], selected_items: List[str] = None) -> Dict[str, List[Dict]]:
        """
        批量解析多个文件
        
        Args:
            file_paths: 文件路径列表
            selected_items: 选中的项目列表
        
        Returns:
            字典，键为文件名，值为解析结果
        """
        results = {}
        
        for file_path in file_paths:
            try:
                filename = os.path.basename(file_path)
                data = self.parse_file(file_path, selected_items)
                results[filename] = data
            except Exception as e:
                results[filename] = [{'项目': '错误', '值': str(e)}]
        
        return results
    
    def export_to_excel(self, data: Dict[str, List[Dict]], output_path: str):
        """
        导出数据到Excel
        
        Args:
            data: 解析结果字典
            output_path: 输出文件路径
        """
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for filename, file_data in data.items():
                # 创建DataFrame
                df = pd.DataFrame(file_data)
                
                # 清理文件名（去掉特殊字符）
                sheet_name = re.sub(r'[\\/*\[\]:?]', '', filename[:31])
                
                # 写入Excel
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # 调整列宽
                worksheet = writer.sheets[sheet_name]
                worksheet.column_dimensions['A'].width = 40
                worksheet.column_dimensions['B'].width = 50
        
        return output_path
