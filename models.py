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


@dataclass
class RoutingGroup:
    group: str                  # PLNNR, e.g. "M0164412"
    popup_row: int              # first row of this group in the "Entry Option: Group" popup
    counters: list[str] = field(default_factory=list)   # GrC / PLNAL, e.g. ["22", "27", "30"]


@dataclass
class RoutingTask:
    material: str
    plant: str = "1901"
    operation: str = "0020"     # operation the new BOM components are assigned to
    excel_rows: list[int] = field(default_factory=list)
    groups: list[RoutingGroup] = field(default_factory=list)
    done: list[str] = field(default_factory=list)       # "M0164412/22" processed OK
    status: str = ""            # SUCCESS / FAILED
    message: str = ""
