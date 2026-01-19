#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PPT to Images Exporter

将 PowerPoint 文件的每个幻灯片导出为图片文件。

支持多种导出方法：
1. LibreOffice 转换（推荐）
2. pdf2image（通过 PDF 中间格式）
3. aspose.slides（功能强大但需要许可证）
"""

import os
import sys
import subprocess
import tempfile
import shutil
from pathlib import Path
from typing import List, Optional, Tuple
from dataclasses import dataclass

# 确保 stdout 使用 UTF-8 编码
if sys.stdout.encoding != 'utf-8':
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')


@dataclass
class ExportConfig:
    """导出配置"""
    dpi: int = 300  # 分辨率
    format: str = "png"  # 输出格式: png, jpg, jpeg
    quality: int = 95  # JPEG 质量 (1-100)
    width: Optional[int] = None  # 自定义宽度（像素）
    height: Optional[int] = None  # 自定义高度（像素）


class PPTExporter:
    """PPT 幻灯片图片导出器"""
    
    def __init__(self, config: Optional[ExportConfig] = None):
        self.config = config or ExportConfig()
        self._check_dependencies()
    
    def _check_dependencies(self):
        """检查必要的依赖"""
        self.has_libreoffice = self._check_libreoffice()
        self.has_pdf2image = self._check_pdf2image()
        self.has_aspose = self._check_aspose()
        
        if not any([self.has_libreoffice, self.has_pdf2image, self.has_aspose]):
            print("⚠️  警告: 未找到任何可用的转换工具")
            print("请安装以下至少一种：")
            print("1. LibreOffice: brew install --cask libreoffice")
            print("2. pdf2image: pip install pdf2image Pillow")
            print("3. aspose.slides: pip install aspose.slides")
    
    def _check_libreoffice(self) -> bool:
        """检查 LibreOffice 是否可用"""
        for cmd in ['soffice', 'libreoffice']:
            if shutil.which(cmd):
                return True
        return False
    
    def _check_pdf2image(self) -> bool:
        """检查 pdf2image 是否可用"""
        try:
            import pdf2image
            from PIL import Image
            return True
        except ImportError:
            return False
    
    def _check_aspose(self) -> bool:
        """检查 aspose.slides 是否可用"""
        try:
            import aspose.slides as slides
            return True
        except ImportError:
            return False
    
    def export(
        self,
        ppt_path: str,
        output_dir: str,
        method: str = "auto",
        prefix: str = "slide",
        progress_callback: Optional[callable] = None
    ) -> List[str]:
        """
        导出 PPT 为图片
        
        Args:
            ppt_path: PPT 文件路径
            output_dir: 输出目录
            method: 导出方法 ('auto', 'libreoffice', 'pdf2image', 'aspose')
            prefix: 输出文件前缀
            progress_callback: 进度回调函数 callback(current, total, filename)
        
        Returns:
            导出的图片文件路径列表
        """
        self.progress_callback = progress_callback
        # 验证输入文件
        if not os.path.exists(ppt_path):
            raise FileNotFoundError(f"文件不存在: {ppt_path}")
        
        if not ppt_path.lower().endswith(('.ppt', '.pptx')):
            raise ValueError("仅支持 .ppt 或 .pptx 格式")
        
        # 创建输出目录
        os.makedirs(output_dir, exist_ok=True)
        
        # 选择导出方法
        if method == "auto":
            if self.has_libreoffice:
                method = "libreoffice"
            elif self.has_pdf2image:
                method = "pdf2image"
            elif self.has_aspose:
                method = "aspose"
            else:
                raise RuntimeError("未找到可用的转换工具")
        
        # 执行导出
        print(f"🚀 使用方法: {method}")
        
        if method == "libreoffice":
            return self._export_with_libreoffice(ppt_path, output_dir, prefix, progress_callback)
        elif method == "pdf2image":
            return self._export_with_pdf2image(ppt_path, output_dir, prefix, progress_callback)
        elif method == "aspose":
            return self._export_with_aspose(ppt_path, output_dir, prefix, progress_callback)
        else:
            raise ValueError(f"不支持的方法: {method}")
    
    def _export_with_libreoffice(
        self,
        ppt_path: str,
        output_dir: str,
        prefix: str,
        progress_callback: Optional[callable] = None
    ) -> List[str]:
        """使用 LibreOffice 导出"""
        print("📄 使用 LibreOffice 转换...")
        
        # 创建临时目录
        with tempfile.TemporaryDirectory() as temp_dir:
            # 第一步：转换为 PDF
            print("  → 转换为 PDF...")
            
            # 设置环境变量以支持中文
            env = os.environ.copy()
            env['LANG'] = 'zh_CN.UTF-8'
            env['LC_ALL'] = 'zh_CN.UTF-8'
            env['SAL_USE_VCLPLUGIN'] = 'svp'  # 使用无头模式
            
            # 检查并设置字体路径
            font_paths = [
                '/usr/share/fonts/truetype/noto',
                '/usr/share/fonts/truetype/wqy',
                '/usr/share/fonts/opentype/noto',
            ]
            
            font_path_str = ':'.join([p for p in font_paths if os.path.exists(p)])
            if font_path_str:
                env['FONTCONFIG_PATH'] = '/etc/fonts'
                env['FONTCONFIG_FILE'] = '/etc/fonts/fonts.conf'
                print(f"  字体路径: {font_path_str}")
            
            cmd = [
                "soffice" if shutil.which("soffice") else "libreoffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", temp_dir,
                ppt_path
            ]
            
            try:
                result = subprocess.run(
                    cmd, 
                    check=True, 
                    capture_output=True, 
                    timeout=120,  # 增加超时时间
                    env=env
                )
                
                # 打印输出以便调试
                if result.stdout:
                    output = result.stdout.decode('utf-8', errors='ignore')
                    if output.strip():
                        print(f"  LibreOffice 输出: {output}")
                
            except subprocess.TimeoutExpired:
                raise RuntimeError("LibreOffice 转换超时（120秒）")
            except subprocess.CalledProcessError as e:
                error_msg = e.stderr.decode('utf-8', errors='ignore')
                raise RuntimeError(f"LibreOffice 转换失败: {error_msg}")
            
            # 找到生成的 PDF
            pdf_files = list(Path(temp_dir).glob("*.pdf"))
            if not pdf_files:
                raise RuntimeError("PDF 转换失败")
            
            pdf_path = str(pdf_files[0])
            
            # 第二步：PDF 转图片
            print("  → PDF 转换为图片...")
            if not self.has_pdf2image:
                raise RuntimeError("需要安装 pdf2image: pip install pdf2image Pillow")
            
            from pdf2image import convert_from_path
            from PIL import Image
            
            images = convert_from_path(
                pdf_path,
                dpi=self.config.dpi,
                fmt=self.config.format
            )
            
            # 保存图片
            output_files = []
            total_images = len(images)
            
            for i, image in enumerate(images, 1):
                # 调整大小（如果指定）
                if self.config.width or self.config.height:
                    image = self._resize_image(image)
                
                # 生成文件名
                filename = f"{prefix}_{i:03d}.{self.config.format}"
                output_path = os.path.join(output_dir, filename)
                
                # 保存
                if self.config.format.lower() in ['jpg', 'jpeg']:
                    image.save(output_path, 'JPEG', quality=self.config.quality)
                else:
                    image.save(output_path, self.config.format.upper())
                
                output_files.append(output_path)
                print(f"  ✓ 已保存: {filename}")
                
                # 调用进度回调
                if progress_callback:
                    progress_callback(i, total_images, filename)
            
            return output_files
    
    def _export_with_pdf2image(
        self,
        ppt_path: str,
        output_dir: str,
        prefix: str,
        progress_callback: Optional[callable] = None
    ) -> List[str]:
        """使用 pdf2image 导出（需要先转 PDF）"""
        print("📄 使用 pdf2image 方法...")
        
        # 这个方法依赖 LibreOffice 先转 PDF
        if not self.has_libreoffice:
            raise RuntimeError("此方法需要 LibreOffice 来转换 PPT 为 PDF")
        
        return self._export_with_libreoffice(ppt_path, output_dir, prefix, progress_callback)
    
    def _export_with_aspose(
        self,
        ppt_path: str,
        output_dir: str,
        prefix: str,
        progress_callback: Optional[callable] = None
    ) -> List[str]:
        """使用 aspose.slides 导出（最强大但需要许可证）"""
        print("📄 使用 Aspose.Slides 转换...")
        
        try:
            import aspose.slides as slides
            from PIL import Image
            import io
        except ImportError:
            raise RuntimeError("需要安装 aspose.slides: pip install aspose.slides")
        
        # 加载演示文稿
        print("  → 加载演示文稿...")
        presentation = slides.Presentation(ppt_path)
        
        output_files = []
        
        try:
            # 遍历每个幻灯片
            for i, slide in enumerate(presentation.slides, 1):
                print(f"  → 导出幻灯片 {i}/{len(presentation.slides)}...")
                
                # 生成缩略图
                scale_x = self.config.width / slide.slide_size.size.width if self.config.width else 2.0
                scale_y = self.config.height / slide.slide_size.size.height if self.config.height else 2.0
                scale = min(scale_x, scale_y) if self.config.width or self.config.height else 2.0
                
                thumbnail = slide.get_thumbnail(scale, scale)
                
                # 转换为 PIL Image
                img_byte_arr = io.BytesIO()
                thumbnail.save(img_byte_arr, slides.ImageFormat.PNG)
                img_byte_arr.seek(0)
                
                pil_image = Image.open(img_byte_arr)
                
                # 生成文件名
                filename = f"{prefix}_{i:03d}.{self.config.format}"
                output_path = os.path.join(output_dir, filename)
                
                # 保存
                if self.config.format.lower() in ['jpg', 'jpeg']:
                    # 转换为 RGB（JPEG 不支持透明度）
                    if pil_image.mode in ('RGBA', 'LA', 'P'):
                        background = Image.new('RGB', pil_image.size, (255, 255, 255))
                        if pil_image.mode == 'P':
                            pil_image = pil_image.convert('RGBA')
                        background.paste(pil_image, mask=pil_image.split()[-1] if pil_image.mode == 'RGBA' else None)
                        pil_image = background
                    
                    pil_image.save(output_path, 'JPEG', quality=self.config.quality)
                else:
                    pil_image.save(output_path, self.config.format.upper())
                
                output_files.append(output_path)
                print(f"  ✓ 已保存: {filename}")
        
        finally:
            presentation.dispose()
        
        return output_files
    
    def _resize_image(self, image):
        """调整图片大小"""
        from PIL import Image
        
        if not self.config.width and not self.config.height:
            return image
        
        orig_width, orig_height = image.size
        
        if self.config.width and self.config.height:
            new_size = (self.config.width, self.config.height)
        elif self.config.width:
            ratio = self.config.width / orig_width
            new_size = (self.config.width, int(orig_height * ratio))
        else:  # self.config.height
            ratio = self.config.height / orig_height
            new_size = (int(orig_width * ratio), self.config.height)
        
        return image.resize(new_size, Image.Resampling.LANCZOS)


def export_ppt_to_images(
    ppt_path: str,
    output_dir: str = "./output",
    method: str = "auto",
    dpi: int = 300,
    format: str = "png",
    prefix: str = "slide"
) -> List[str]:
    """
    便捷函数：导出 PPT 为图片
    
    Args:
        ppt_path: PPT 文件路径
        output_dir: 输出目录
        method: 导出方法 ('auto', 'libreoffice', 'pdf2image', 'aspose')
        dpi: 分辨率
        format: 输出格式 ('png', 'jpg', 'jpeg')
        prefix: 文件前缀
    
    Returns:
        导出的图片文件路径列表
    
    Example:
        >>> files = export_ppt_to_images('presentation.pptx', './images')
        >>> print(f"导出了 {len(files)} 张图片")
    """
    config = ExportConfig(dpi=dpi, format=format)
    exporter = PPTExporter(config)
    return exporter.export(ppt_path, output_dir, method=method, prefix=prefix)


if __name__ == "__main__":
    # 简单的命令行测试
    if len(sys.argv) < 2:
        print("用法: python ppt_exporter.py <ppt文件路径> [输出目录]")
        print("示例: python ppt_exporter.py presentation.pptx ./images")
        sys.exit(1)
    
    ppt_file = sys.argv[1]
    output_directory = sys.argv[2] if len(sys.argv) > 2 else "./output"
    
    try:
        files = export_ppt_to_images(ppt_file, output_directory)
        print(f"\n✅ 成功导出 {len(files)} 张图片到 {output_directory}")
    except Exception as e:
        print(f"\n❌ 错误: {e}")
        sys.exit(1)
