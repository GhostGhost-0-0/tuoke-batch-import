"""
文件操作处理模块
"""

import os
import glob


class FileHandler:
    """文件操作处理类"""

    def resolve_links_file_path(
        self, site: str, collection_word: str, links_dir: str = "links"
    ) -> str:
        """
        解析与 load 规则一致的链接文件路径（多文件时取修改时间最新）。
        找不到返回空字符串。
        """
        site_dir = os.path.join(links_dir, site)
        if not os.path.exists(site_dir):
            return ""

        pattern_with_suffix = os.path.join(site_dir, f"{collection_word}_*.txt")
        exact_file = os.path.join(site_dir, f"{collection_word}.txt")

        matching_files = glob.glob(pattern_with_suffix)
        if os.path.exists(exact_file):
            matching_files.append(exact_file)

        if not matching_files:
            return ""

        return max(matching_files, key=os.path.getmtime)

    def write_links_to_file(
        self,
        site: str,
        collection_word: str,
        content: str,
        links_dir: str = "links",
    ) -> bool:
        """
        将链接文本写回当前采集词对应的 txt（与读取时为同一文件）。
        """
        link_file = self.resolve_links_file_path(site, collection_word, links_dir)
        if not link_file:
            print(
                f"  ⚠️ 无法写回链接文件: 未找到 {collection_word}_*.txt 或 {collection_word}.txt"
            )
            return False
        try:
            text = (content or "").rstrip("\n\r")
            with open(link_file, "w", encoding="utf-8") as f:
                f.write(text + ("\n" if text else ""))
            print(f"  ✓ 已写回链接文件: {os.path.basename(link_file)}")
            return True
        except Exception as e:
            print(f"  ❌ 写回链接文件失败: {e}")
            return False
    
    def load_links_from_file(self, site: str, collection_word: str, links_dir: str = "links") -> str:
        """
        从links目录读取链接数据
        :param site: 站点名称（如"台湾"、"新加坡"等）
        :param collection_word: 采集词
        :param links_dir: links目录路径
        :return: 链接数据（换行分隔的字符串）
        """
        site_dir = os.path.join(links_dir, site)
        
        if not os.path.exists(site_dir):
            print(f"  ⚠️ 站点目录不存在: {site_dir}")
            return ""

        link_file = self.resolve_links_file_path(site, collection_word, links_dir)
        if not link_file:
            print(f"  ⚠️ 未找到匹配的链接文件: {collection_word}_*.txt 或 {collection_word}.txt")
            return ""

        print(f"  ✓ 找到链接文件: {os.path.basename(link_file)}")
        
        try:
            with open(link_file, "r", encoding="utf-8") as f:
                links = f.read().strip()
            
            # 统计链接数量
            link_count = len([line for line in links.split('\n') if line.strip()])
            print(f"  ✓ 读取到 {link_count} 个链接")
            
            return links
        except Exception as e:
            print(f"  ❌ 读取链接文件失败: {e}")
            return ""

