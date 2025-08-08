"""
Word document processing utilities.
This module is deprecated and will be removed in a future version.
Please use precisiondoc.utils.word.WordUtils instead.
"""
import warnings
from .word.document_formatting import DocumentFormatter
from .word.table_utils import TableUtils
from .word.content_formatting import ContentFormatter
from .word.evidence_processing import EvidenceProcessor
from .word.image_utils import ImageUtils
from .word.export_utils import ExportUtils
from .word import WordUtils

# 发出弃用警告
warnings.warn(
    "The WordUtils class in word_utils.py is deprecated. "
    "Please use precisiondoc.utils.word.WordUtils instead.",
    DeprecationWarning,
    stacklevel=2
)

# 为了保持向后兼容性，重新导出WordUtils类
__all__ = ['WordUtils']
