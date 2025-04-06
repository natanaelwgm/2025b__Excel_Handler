# Make key functions available directly from the 'excel_handler.gemini' package
from .excel_handler_gemini import (
    read_excel_file_data,
    compare_excel_data,
    create_demo_excel_file,
    main,
    write_excel_data_to_txt,
    write_comparison_summary_to_txt
)

# Define __all__ to specify the public API
__all__ = [
    'read_excel_file_data',
    'compare_excel_data',
    'create_demo_excel_file',
    'main',
    'write_excel_data_to_txt',
    'write_comparison_summary_to_txt'
] 