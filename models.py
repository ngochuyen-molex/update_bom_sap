from dataclasses import dataclass, field


@dataclass
class BomItem:
    excel_row: int          # Excel row to write the result back to
    component: str
    quantity: float
    unit: str
    scrap: float = 0.0
    item_category: str = "L"


@dataclass
class BomHeader:
    parent: str
    alt_bom: str
    base_qty: float
    plant: str = "1901"
    usage: str = "1"
    action: str = "CREATE"      # CREATE (CS01) or CHANGE (CS02)
    items: list[BomItem] = field(default_factory=list)
    created_alt_bom: str = ""   # alt BOM number SAP assigned in CS01
    status: str = ""        # SUCCESS / FAILED
    message: str = ""

    @property
    def excel_rows(self):
        return [item.excel_row for item in self.items]
