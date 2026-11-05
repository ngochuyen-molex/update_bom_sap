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
class Routing:
    material: str
    group: str                  # PLNNR, e.g. "M0164412"
    counter: str                # GrC / PLNAL without leading zeros, e.g. "22"
    description: str = ""       # KTEXT
    plant: str = "1901"
    group_size: int = 1         # number of GrC of this material + group (from Snowflake)
    status: str = ""            # SUCCESS / SKIPPED / FAILED
    message: str = ""

    @property
    def key(self):
        return f"{self.material} {self.group}/{self.counter}"


@dataclass
class ProductionVersion:
    material: str
    group: str                  # routing group (PLNNR)
    counter: str                # group counter (PLNAL)
    description: str            # routing description
    alt_bom: str
    bom_usage: str = "1"
    text: str = ""              # production version text = description + " B" + alt BOM
    version: str = ""           # new version SAP gets, e.g. "0010"
    status: str = ""            # SUCCESS / FAILED
    message: str = ""

    @property
    def key(self):
        return f"{self.material} {self.group}/{self.counter}"
