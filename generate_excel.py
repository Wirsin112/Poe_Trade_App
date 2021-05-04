import xlsxwriter
import yaml
from recipe import Recipe
yaml.warnings({'YAMLLoadWarning': False})


def generate_excel():
    # Read all recipes
    with open('recipes.yaml', 'r') as f:
        yaml_file = yaml.load(f)

    # Create workbook and worksheets
    workbook = xlsxwriter.Workbook('output.xlsx')
    worksheet_vendor = workbook.add_worksheet('Vendor Recipes')

    ######################################################################
    #########################   Vendor Recipes   #########################
    ######################################################################

    # Set columns and cells
    worksheet_vendor.set_column("A:M", 20)
    worksheet_vendor.set_row(0, 45)
    worksheet_vendor.set_row(1, 25)

    # Title
    title_format = workbook.add_format({
        "bg_color": "#B1BCBE",
        "font": "Century",
        "font_size": 22,
        "bold": True,
        "border": 1,
        "border_color": "#000000"
    })
    title_format.set_align('center')
    title_format.set_align('vcenter')
    worksheet_vendor.merge_range("A1:M1", "Vendor Recipes", title_format)

    # Headers
    header_format = workbook.add_format({
        "bg_color": "#F5F5EC",
        "font": "Arial",
        "font_size": 12,
        "bold": True,
        "border": 1,
        "border_color": "#000000"
    })
    header_format.set_align('center')
    header_format.set_align('vcenter')
    worksheet_vendor.merge_range("A2:H2", "Components", header_format)
    worksheet_vendor.merge_range("I2:J2", "Result", header_format)
    worksheet_vendor.write("K2", "ROI", header_format)
    worksheet_vendor.write("L2", "Profit", header_format)
    worksheet_vendor.write("M2", "Wiki links", header_format)

    # Load all recipes
    recipes = []
    for recipe_yaml in yaml_file['vendor_recipes']:
        current_recipe = yaml_file['vendor_recipes'][recipe_yaml]
        recipes.append(Recipe(recipe_yaml, 'Ritual', current_recipe['components'], current_recipe['results'], current_recipe['wiki']))

    # Sort recipes descending by profit
    recipes.sort(key=lambda x: x.profit, reverse=True)

    # Items name formatting based on liquidity of the item
    liquidity_color_palette = {0: "#FF6962",
                               1: "#FFB6B3",
                               2: "#FFD5D4",
                               3: "#E7F1E8",
                               4: "#BDE7BD",
                               5: "#77DD76"}
    items_formats = []
    for liquidity in range(6):
        items_formats.append(workbook.add_format({"font": "Arial",
                                                  "font_size": 11,
                                                  "bg_color": liquidity_color_palette[liquidity]
                                                  }))
        items_formats[liquidity].set_align('center')
        items_formats[liquidity].set_align('vcenter')

    # Numbers formatting
    numbers_format = workbook.add_format({"font": "Calibri",
                                          "font_size": 11,
                                          "bottom": 1,
                                          "bottom_color": "#000000"})
    numbers_format.set_align('center')
    numbers_format.set_align('vcenter')

    # Add items to the worksheet
    row = 2
    for recipe in recipes:
        # Components
        col = 0
        # Draw border between items
        for i in range(10):
            worksheet_vendor.write(row + 1, i, "", numbers_format)

        for component_item in recipe.components:
            # Name of the item and url
            item = component_item[0]
            count = component_item[1]

            worksheet_vendor.merge_range(row, col, row, col + 1, "placeholder")

            # Set bg_color based on liquidity
            worksheet_vendor.write_url(row, col, item.search_link, items_formats[item.liquidity], item.name)

            # Price and count
            worksheet_vendor.write(row + 1, col, f"{item.price}c", numbers_format)
            worksheet_vendor.write(row + 1, col + 1, f"x {count}", numbers_format)
            col += 2

        # Result
        col = 8
        item = recipe.results[0][0]
        count = recipe.results[0][1]

        worksheet_vendor.merge_range(row, col, row, col + 1, "placeholder")

        # Set bg_color based on liquidity
        worksheet_vendor.write_url(row, col, item.search_link, items_formats[item.liquidity], item.name)

        # Price and count
        worksheet_vendor.write(row + 1, col, f"{item.price}c", numbers_format)
        worksheet_vendor.write(row + 1, col + 1, f"x {count}", numbers_format)

        # ROI
        col = 10
        worksheet_vendor.merge_range(row, col, row + 1, col, f"{recipe.roi} %", numbers_format)

        # Profit
        col = 11
        worksheet_vendor.merge_range(row, col, row + 1, col, recipe.profit, numbers_format)

        # Wiki
        col = 12
        worksheet_vendor.merge_range(row, col, row + 1, col, "placeholder", numbers_format)
        worksheet_vendor.write_url(row, col, recipe.wiki, numbers_format, "Wiki Link")

        # Start next row
        row += 2

    ######################################################################
    #######################   Harbinger Upgrades   #######################
    ######################################################################

    # Create worksheet
    worksheet_harbinger = workbook.add_worksheet("Harbinger Upgrades")

    # Format columns and rows
    worksheet_harbinger.set_column("A:I", 20)
    worksheet_harbinger.set_row(0, 45)
    worksheet_harbinger.set_row(1, 25)

    # Title
    worksheet_harbinger.merge_range("A1:I1", "Harbinger Upgrades", title_format)

    # Headers
    worksheet_harbinger.merge_range("A2:B2", "Item", header_format)
    worksheet_harbinger.merge_range("C2:D2", "Scroll", header_format)
    worksheet_harbinger.merge_range("E2:F2", "Result", header_format)
    worksheet_harbinger.write("G2", "ROI", header_format)
    worksheet_harbinger.write("H2", "Profit", header_format)
    worksheet_harbinger.write("I2", "Wiki links", header_format)

    # Load all recipes
    recipes = []
    for recipe_yaml in yaml_file['harbinger_upgrades']:
        current_recipe = yaml_file['harbinger_upgrades'][recipe_yaml]
        recipes.append(Recipe(recipe_yaml, 'Ritual', current_recipe['components'], current_recipe['results'], current_recipe['wiki']))

    # Sort recipes descending by profit
    recipes.sort(key=lambda x: x.profit, reverse=True)

    # Add items to the worksheet
    row = 2
    for recipe in recipes:
        # Components
        col = 0
        # Draw border between items
        for i in range(9):
            worksheet_harbinger.write(row + 1, i, "", numbers_format)

        for component_item in recipe.components:
            # Name of the item and url
            item = component_item[0]
            count = component_item[1]

            worksheet_harbinger.merge_range(row, col, row, col + 1, "placeholder")

            # Set bg_color based on liquidity
            worksheet_harbinger.write_url(row, col, item.search_link, items_formats[item.liquidity], item.name)

            # Price and count
            worksheet_harbinger.write(row + 1, col, f"{item.price}c", numbers_format)
            worksheet_harbinger.write(row + 1, col + 1, f"x {count}", numbers_format)
            col += 2

        # Result
        col = 4
        item = recipe.results[0][0]
        count = recipe.results[0][1]

        worksheet_harbinger.merge_range(row, col, row, col + 1, "placeholder")

        # Set bg_color based on liquidity
        worksheet_harbinger.write_url(row, col, item.search_link, items_formats[item.liquidity], item.name)

        # Price and count
        worksheet_harbinger.write(row + 1, col, f"{item.price}c", numbers_format)
        worksheet_harbinger.write(row + 1, col + 1, f"x {count}", numbers_format)

        # ROI
        col = 6
        worksheet_harbinger.merge_range(row, col, row + 1, col, f"{recipe.roi} %", numbers_format)

        # Profit
        col = 7
        worksheet_harbinger.merge_range(row, col, row + 1, col, recipe.profit, numbers_format)

        # Wiki
        col = 8
        worksheet_harbinger.merge_range(row, col, row + 1, col, "placeholder", numbers_format)
        worksheet_harbinger.write_url(row, col, recipe.wiki, numbers_format, "Wiki Link")

        # Start next row
        row += 2

    ######################################################################
    ##########################   Vial Uniques   ##########################
    ######################################################################

    # Create worksheet
    worksheet_vials = workbook.add_worksheet("Vial Uniques")

    # Format columns and rows
    worksheet_vials.set_column("A:I", 20)
    worksheet_vials.set_row(0, 45)
    worksheet_vials.set_row(1, 25)

    # Title
    worksheet_vials.merge_range("A1:I1", "Vial Uniques", title_format)

    # Headers
    worksheet_vials.merge_range("A2:B2", "Item", header_format)
    worksheet_vials.merge_range("C2:D2", "Vial", header_format)
    worksheet_vials.merge_range("E2:F2", "Result", header_format)
    worksheet_vials.write("G2", "ROI", header_format)
    worksheet_vials.write("H2", "Profit", header_format)
    worksheet_vials.write("I2", "Wiki links", header_format)

    # Load all recipes
    recipes = []
    for recipe_yaml in yaml_file['vial_uniques']:
        current_recipe = yaml_file['vial_uniques'][recipe_yaml]
        recipes.append(Recipe(recipe_yaml, 'Ritual', current_recipe['components'], current_recipe['results'], current_recipe['wiki']))

    # Sort recipes descending by profit
    recipes.sort(key=lambda x: x.profit, reverse=True)

    # Add items to the worksheet
    row = 2
    for recipe in recipes:
        # Components
        col = 0
        # Draw border between items
        for i in range(9):
            worksheet_vials.write(row + 1, i, "", numbers_format)

        for component_item in recipe.components:
            # Name of the item and url
            item = component_item[0]
            count = component_item[1]

            worksheet_vials.merge_range(row, col, row, col + 1, "placeholder")

            # Set bg_color based on liquidity
            worksheet_vials.write_url(row, col, item.search_link, items_formats[item.liquidity], item.name)

            # Price and count
            worksheet_vials.write(row + 1, col, f"{item.price}c", numbers_format)
            worksheet_vials.write(row + 1, col + 1, f"x {count}", numbers_format)
            col += 2

        # Result
        col = 4
        item = recipe.results[0][0]
        count = recipe.results[0][1]

        worksheet_vials.merge_range(row, col, row, col + 1, "placeholder")

        # Set bg_color based on liquidity
        worksheet_vials.write_url(row, col, item.search_link, items_formats[item.liquidity], item.name)

        # Price and count
        worksheet_vials.write(row + 1, col, f"{item.price}c", numbers_format)
        worksheet_vials.write(row + 1, col + 1, f"x {count}", numbers_format)

        # ROI
        col = 6
        worksheet_vials.merge_range(row, col, row + 1, col, f"{recipe.roi} %", numbers_format)

        # Profit
        col = 7
        worksheet_vials.merge_range(row, col, row + 1, col, recipe.profit, numbers_format)

        # Wiki
        col = 8
        worksheet_vials.merge_range(row, col, row + 1, col, "placeholder", numbers_format)
        worksheet_vials.write_url(row, col, recipe.wiki, numbers_format, "Wiki Link")

        # Start next row
        row += 2

    ######################################################################
    #######################   Blessing Upgrades   ########################
    ######################################################################

    # Create worksheet
    worksheet_blessings = workbook.add_worksheet("Blessing Upgrades")

    # Format columns and rows
    worksheet_blessings.set_column("A:I", 20)
    worksheet_blessings.set_row(0, 45)
    worksheet_blessings.set_row(1, 25)

    # Title
    worksheet_blessings.merge_range("A1:I1", "Blessing Upgrades", title_format)

    # Headers
    worksheet_blessings.merge_range("A2:B2", "Item", header_format)
    worksheet_blessings.merge_range("C2:D2", "Blessing", header_format)
    worksheet_blessings.merge_range("E2:F2", "Result", header_format)
    worksheet_blessings.write("G2", "ROI", header_format)
    worksheet_blessings.write("H2", "Profit", header_format)
    worksheet_blessings.write("I2", "Wiki links", header_format)

    # Load all recipes
    recipes = []
    for recipe_yaml in yaml_file['blessing_upgrades']:
        current_recipe = yaml_file['blessing_upgrades'][recipe_yaml]
        recipes.append(Recipe(recipe_yaml, 'Ritual', current_recipe['components'], current_recipe['results'], current_recipe['wiki']))

    # Sort recipes descending by profit
    recipes.sort(key=lambda x: x.profit, reverse=True)

    # Add items to the worksheet
    row = 2
    for recipe in recipes:
        # Components
        col = 0
        # Draw border between items
        for i in range(9):
            worksheet_blessings.write(row + 1, i, "", numbers_format)

        for component_item in recipe.components:
            # Name of the item and url
            item = component_item[0]
            count = component_item[1]

            worksheet_blessings.merge_range(row, col, row, col + 1, "placeholder")

            # Set bg_color based on liquidity
            worksheet_blessings.write_url(row, col, item.search_link, items_formats[item.liquidity], item.name)

            # Price and count
            worksheet_blessings.write(row + 1, col, f"{item.price}c", numbers_format)
            worksheet_blessings.write(row + 1, col + 1, f"x {count}", numbers_format)
            col += 2

        # Result
        col = 4
        item = recipe.results[0][0]
        count = recipe.results[0][1]

        worksheet_blessings.merge_range(row, col, row, col + 1, "placeholder")

        # Set bg_color based on liquidity
        worksheet_blessings.write_url(row, col, item.search_link, items_formats[item.liquidity], item.name)

        # Price and count
        worksheet_blessings.write(row + 1, col, f"{item.price}c", numbers_format)
        worksheet_blessings.write(row + 1, col + 1, f"x {count}", numbers_format)

        # ROI
        col = 6
        worksheet_blessings.merge_range(row, col, row + 1, col, f"{recipe.roi} %", numbers_format)

        # Profit
        col = 7
        worksheet_blessings.merge_range(row, col, row + 1, col, recipe.profit, numbers_format)

        # Wiki
        col = 8
        worksheet_blessings.merge_range(row, col, row + 1, col, "placeholder", numbers_format)
        worksheet_blessings.write_url(row, col, recipe.wiki, numbers_format, "Wiki Link")

        # Start next row
        row += 2

    workbook.close()


if __name__ == '__main__':
    try:
        generate_excel()
        print('Workbook was updated!')
    except xlsxwriter.exceptions.FileCreateError:
        print('Cannot update excel - close the workbook!')
