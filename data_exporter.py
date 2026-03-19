import pandas as pd
import json
import xml.etree.ElementTree as ET
from xml.dom import minidom
import os

OUTPUT_DIR = "renteasy_data/"


def save_csv(data: list, filename: str):
    """Save via pandas, UTF-8-sig encoding (opens correctly in Excel)"""
    df = pd.DataFrame(data)
    df.to_csv(OUTPUT_DIR + filename, index=False, encoding='utf-8-sig')
    print(f"CSV сохранён: {filename} ({len(df)} строк)")


def save_json(data: list, filename: str):
    """Save pretty-printed JSON with metadata"""
    with open(OUTPUT_DIR + filename, 'w', encoding='utf-8') as f:
        json.dump({
            "metadata": {
                "title": "RentEasy KZ — Rental Market Dataset",
                "description": "Объявления аренды квартир в Казахстане",
                "total_records": len(data),
                "cities": ["Астана", "Алматы", "Шымкент", "Карагандa", "Актобе"],
                "sources": ["krisha.kz", "olx.kz"]
            },
            "listings": data
        }, f, ensure_ascii=False, indent=2)
    print(f"JSON сохранён: {filename}")


def save_xml(data: list, filename: str):
    """Generate readable XML with metadata"""
    root = ET.Element("renteasy_dataset")

    meta = ET.SubElement(root, "metadata")
    ET.SubElement(meta, "title").text = "RentEasy KZ — Rental Market Dataset"
    ET.SubElement(meta, "description").text = "Объявления аренды квартир в Казахстане"
    ET.SubElement(meta, "total_records").text = str(len(data))
    ET.SubElement(meta, "sources").text = "krisha.kz, olx.kz"

    listings_el = ET.SubElement(root, "listings")
    for item in data:
        listing = ET.SubElement(listings_el, "listing")
        listing.set("id", str(item.get("id", "")))
        for key, val in item.items():
            if key == "id":
                continue
            el = ET.SubElement(listing, key)
            el.text = str(val) if val is not None else ""

    xml_str = minidom.parseString(ET.tostring(root, encoding='unicode')).toprettyxml(indent="  ")
    with open(OUTPUT_DIR + filename, 'w', encoding='utf-8') as f:
        f.write(xml_str)
    print(f"XML сохранён: {filename}")


if __name__ == "__main__":
    from parser import main as parse_all
    data = parse_all()

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    save_csv(data, "renteasy_dataset.csv")
    save_json(data, "renteasy_dataset.json")
    save_xml(data, "renteasy_dataset.xml")
