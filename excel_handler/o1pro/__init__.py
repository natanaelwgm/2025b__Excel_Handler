# Make key functions available directly from the 'excel_handler.o1pro' package
from .excel_handler_o1pro import (
    read_excel_to_dict,
    compare_excel_dicts,
    export_excel_dict_to_txt,
    export_comparison_to_txt,
    create_dummy_excel_files,
    main
)

# Define the public API
__all__ = [
    'read_excel_to_dict',
    'compare_excel_dicts',
    'export_excel_dict_to_txt',
    'export_comparison_to_txt',
    'create_dummy_excel_files',
    'main'
] 