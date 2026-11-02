from bom import clean_component, set_text, press_enter

# ============================================================
# NAVIGATION
# ============================================================
def open_ca02(session):
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NCA02"
    press_enter(session)

def open_routing(session, material, plant="1901"):
    print(f"Material={clean_component(material)}, Plant={plant}")
    set_text(session, "wnd[0]/usr/ctxt[0]", clean_component(material))
    set_text(session, "wnd[0]/usr/ctxt[1]", plant)

    press_enter(session)
