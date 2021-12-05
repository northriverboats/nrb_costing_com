#!/usr/bin/env python
# vim expandtab shiftwidth=4 softtabstop=4
# pylint: disable=too-many-lines
"""
Costing Sheets Merge Boat and Cabin BOMs
"""
from .boms import Bom, BomSection, MergedBom, MergedPart, MergedSection
from .models import Model
from .utilities import logger

def merge_sections(boat_sections: list[BomSection],
                   size: str) -> list[MergedSection]:
    """merge all sections by filtering out parts and applying qty
    adujustments

    Arguments:

    Returns:
        sections -- list of the resutling merged sections

    """
    merged_sections: list[MergedSection] = []
    if not boat_sections:
        return merged_sections
    for section in boat_sections:
        merged_section = MergedSection(section.name, {})
        for part_number in section.parts:
            bom_parts = section.parts[part_number]
            if part_number not in merged_section.parts:
                merged_section.parts[part_number] = MergedPart(
                    0.0,
                    bom_parts[0].description,
                    bom_parts[0].uom,
                    bom_parts[0].unitprice,
                    bom_parts[0].vendor,
                    bom_parts[0].updated,
                    0.0
                )
            merged_part = merged_section.parts[part_number]
            for bom_part in bom_parts:
                if (bom_part.smallest > 0 and not
                    bom_part.smallest <=
                    float(size) <=
                    bom_part.biggest):
                    continue
                factor = (1 if not bom_part.percent
                          else float(size) / bom_part.percent)
                merged_part.qty = merged_part.qty + bom_part.qty  * factor
                merged_part.total = merged_part.qty * merged_part.unitprice
        merged_section.total = sum([merged_section.parts[key].total
                                    for key in merged_section.parts])
        merged_sections.append(merged_section)
    return merged_sections

def merge_labor(boat_bom: Bom, cabin_bom: Bom, size: str) -> dict[str, float]:
    """combine labor for boat and cabin

    Arguments:
        boat_bom -- bom with sizes/labors
        cabin_bom  -- bom with sizes/labor, "0" size = any
        size -- size of boat to use

    Returns:
        boat_labor -- combined labor hours
    """
    boat_labor = boat_bom.sizes[size]
    cabin_labor = cabin_bom.sizes.get("0", {})
    if size in cabin_bom.sizes:
        cabin_labor = cabin_bom.sizes[size]

    for dept, hours in cabin_labor.items():
        boat_labor[dept] += hours
    return boat_labor

def combine_sections(boat_sections: list[MergedSection],
                     cabin_sections: list[MergedSection]
                    ) -> list[MergedSection]:
    """combine sections"""
    if not cabin_sections:
        return boat_sections
    for boat_section, cabin_section in zip(boat_sections, cabin_sections):
        boat_parts = boat_section.parts
        cabin_parts = cabin_section.parts
        boat_section.total += cabin_section.total
        for part_number in cabin_parts:
            if part_number in  boat_parts:
                boat_parts[part_number].qty += cabin_parts[part_number].qty
                boat_parts[part_number].total += cabin_parts[part_number].total
            else:
                boat_parts[part_number] = cabin_parts[part_number]
    return boat_sections


def merge_boms(boat_bom: Bom, cabin_bom: Bom, size: str) -> MergedBom:
    """Merge bom and hours

    Arguments:
        boat_bom -- boat size/labor parts
        cabin_bom - cabin size/labor parts
        size -- size of boat we want to create MergedBom for

    Returns:
        MergedBom
    """
    boat_sections: list[MergedSection]  = merge_sections(boat_bom.sections,
                                                         size)
    cabin_sections: list[MergedSection]  = merge_sections(cabin_bom.sections,
                                                          size)
    sections: list[MergedSection] = combine_sections(boat_sections,
                                                     cabin_sections)
    labor = merge_labor(boat_bom, cabin_bom, size)
    return MergedBom(boat_bom.name, boat_bom.beam, size, labor, sections)

def get_bom(boms: dict[str, Bom], model: Model, size: str) -> Bom:
    """Merges sheets if necessary and returns a BOM.
    Assumes if sheet is not None that there will be a match

    Arguments:
        bom: list[Bom] --
        model: Model -- sheet1 can not be None and must be found
                        sheet2 can be None but *must* be found if not None

    Returns:
        Bom -- Returns new Bom of combined Bom(s)
    """
    boat_bom: Bom = (boms[model.sheet1]
                     if model.sheet1 in boms
                     else  Bom('', "", 0.0, 0.0, {}, []))
    cabin_bom: Bom = (boms[model.sheet2]
                      if model.sheet2 in boms
                      else  Bom('', "", 0.0, 0.0, {}, []))
    if boat_bom.name == "":
        logger.debug("boat_bom not found error %s", model.sheet1)
    return merge_boms(boat_bom, cabin_bom, size)
