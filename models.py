from dataclasses import dataclass, field
from config import BOM_USAGE, PLANT


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
    plant: str = PLANT
    usage: str = BOM_USAGE
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
    plant: str = PLANT
    group_size: int = 1         # number of GrC of this material + group (from Snowflake)
    deleted: bool = False       # routing header deletion indicator (PLKO.DELKZ = X)
    # Component assignment per routing x alt BOM (empty alt_bom = not used, e.g. production version)
    alt_bom: str = ""           # alt BOM without leading zeros, e.g. "2"
    alt_bom_text: str = ""      # STKO.STKTX of the latest change state
    bom_count: int = 1          # active alt BOMs of the material (BOM popup check)
    op_0030_wc: str = ""        # work center of operation 0030 (lubricant rule)
    unassigned: int = -1        # unassigned components (Snowflake); -1 = unknown
    total_items: int = -1       # BOM items of this alt BOM (Snowflake); -1 = unknown
    status: str = ""            # SUCCESS / SKIPPED / FAILED / PENDING
    message: str = ""

    @property
    def key(self):
        alt = f" alt {self.alt_bom}" if self.alt_bom else ""
        return f"{self.material} {self.group}/{self.counter}{alt}"


@dataclass
class ProductionVersion:
    material: str
    group: str                  # routing group (PLNNR)
    counter: str                # group counter (PLNAL)
    description: str            # routing description
    alt_bom: str
    bom_usage: str = BOM_USAGE
    text: str = ""              # production version text = description + " B" + alt BOM
    version: str = ""           # new version SAP gets, e.g. "0010"
    status: str = ""            # SUCCESS / FAILED
    message: str = ""

    @property
    def key(self):
        return f"{self.material} {self.group}/{self.counter}"
