"""
配置管理模块
负责加载和管理配置文件
"""

import json
import os
from typing import Dict, Any


class ConfigManager:
    """配置管理器"""
    
    DEFAULT_CONFIG = {
        "类目": {
            "一级": "美妆保健(美妆保健)",
            "二级": "手足保养与美甲(手足保养与美甲)",
            "三级": "手部保养(手部保养)",
            "四级": "手膜"
        },
        "属性": {
            "品牌": "chloe",
            "保质期": "24个月"
        }
    }
    
    def __init__(self, config_file: str = "属性配置.json"):
        """
        初始化配置管理器
        :param config_file: 配置文件路径
        """
        self.config_file = config_file
        self._config = None
    
    def load(self) -> Dict[str, Any]:
        """
        加载配置文件
        :return: 配置字典
        """
        if self._config is not None:
            return self._config
        
        if not os.path.exists(self.config_file):
            print(f"⚠️ 配置文件 '{self.config_file}' 不存在，使用默认配置")
            self._config = self.DEFAULT_CONFIG.copy()
            return self._config
        
        try:
            with open(self.config_file, "r", encoding="utf-8") as f:
                config = json.load(f)
            
            # 兼容旧格式（如果直接是属性配置）
            if "类目" not in config and "属性" not in config:
                print(f"ℹ️ 检测到旧格式配置文件，转换为新格式")
                config = {
                    "类目": self.DEFAULT_CONFIG["类目"].copy(),
                    "属性": config
                }
            
            # 确保类目和属性都存在
            if "类目" not in config:
                config["类目"] = self.DEFAULT_CONFIG["类目"].copy()
            if "属性" not in config:
                config["属性"] = {}
            
            print(f"✅ 已从 '{self.config_file}' 加载配置")
            print(f"   类目: {config.get('类目', {})}")
            print(f"   属性: {config.get('属性', {})}")
            
            self._config = config
            return config
        except Exception as e:
            print(f"❌ 读取配置文件失败: {e}")
            self._config = {
                "类目": self.DEFAULT_CONFIG["类目"].copy(),
                "属性": {}
            }
            return self._config
    
    def get_category_config(self) -> Dict[str, str]:
        """获取类目配置"""
        config = self.load()
        return config.get("类目", {})
    
    def get_attribute_config(self) -> Dict[str, str]:
        """获取属性配置"""
        config = self.load()
        return config.get("属性", {})
    
    def reload(self) -> Dict[str, Any]:
        """重新加载配置"""
        self._config = None
        return self.load()

