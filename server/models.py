"""
models.py - Pydantic 데이터 모델
"""
from pydantic import BaseModel
from typing import Optional, List


class IncomingDataRow(BaseModel):
    id: str
    company_code: str = ""
    company_name: str
    date: str
    time: str = ""
    tm_no: str
    product_name: str
    quantity: int = 0
    pdf_url: str = ""
    created_at: str = ""
    created_by: str = ""
    updated_at: str = ""


class InspectionResultRow(BaseModel):
    id: str
    company_code: str = ""
    date: str
    company_name: str
    tm_no: str
    product_name: str
    inspection_item: str
    inspection_type: str = "정량"
    measurement_method: str = ""
    lower_limit: str = ""
    upper_limit: str = ""
    samples: List[str] = []
    pass_fail_result: str = ""
    registered_at: str = ""
    registered_by: str = ""


class ItemListRow(BaseModel):
    company_code: str = ""
    tm_no: str
    company_name: str
    product_name: str = ""
    inspection_type: str = "검사"


class SyncBatchRequest(BaseModel):
    incoming_data: List[IncomingDataRow] = []
    inspection_results: List[InspectionResultRow] = []
    item_list: List[ItemListRow] = []


class PdfSyncRequest(BaseModel):
    company_name: str
    year: str
    month: str
    file_name: str
    file_data: str          # base64 encoded
    google_file_id: str = ""
    tm_no: str = ""


class AppConfigRequest(BaseModel):
    sheet_id:  str = ""
    folder_id: str = ""


class HistoryFilter(BaseModel):
    company_name: Optional[str] = None
    date_from: Optional[str] = None
    date_to: Optional[str] = None
    tm_no: Optional[str] = None
    inspection_type: Optional[str] = None
    page: int = 1
    page_size: int = 50
