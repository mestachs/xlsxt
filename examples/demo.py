import xlsxt 
import json
import string
import random
import sys

# use to load test a bit the solution
generate_test_data = len(sys.argv) == 2 and sys.argv[1] == "loadtest"

if generate_test_data:

    def generate_fosas(region_index, district_letter, count):
        fosas = []
        for i in range(count):
            fosa_name = (
                f"fosa {region_index}{district_letter}{string.ascii_uppercase[i]} CS"
            )
            quantity = (i % 5) + 1  # Distribute quantities randomly
            fosas.append({"name": fosa_name, "quantity": quantity})
        return fosas

    def generate_districts(region_index, district_count, fosas_per_district):
        districts = []
        for i in range(district_count):
            district_letter = string.ascii_uppercase[i]
            name = f"District {region_index}.{district_letter}"
            settlement = "urban" if i % 2 == 0 else "rural"
            quantity = (i + 1) * 5
            fosa_count = fosas_per_district + random.randint(-2, 2)
            fosas = generate_fosas(
                region_index, district_letter, fosa_count
            )  # Vary fosa count

            districts.append(
                {
                    "name": name,
                    "settlement": settlement,
                    "quantity": quantity,
                    "fosas": fosas,
                }
            )
        return districts

    def generate_regions(region_count, districts_per_region, fosas_per_district):
        regions = []
        for i in range(region_count):
            name = f"Region {i + 1}"
            region_index = i + 1
            district_variation = districts_per_region + random.randint(-2, 2)
            districts = generate_districts(
                region_index, district_variation, fosas_per_district
            )
            total = sum(d["quantity"] for d in districts)

            region_data = {"name": name, "districts": districts, "total": total}
            regions.append(region_data)
        return regions

    def generate_test_data(
        region_count=7, districts_per_region=10, fosas_per_district=10
    ):
        return {
            "name": "Ministère",
            "sample_url": "https://google.com",
            "regions": generate_regions(
                region_count, districts_per_region, fosas_per_district
            ),
        }

    context = generate_test_data()
else:
    with open("./examples/demo.json", "r", encoding="utf-8") as file:
        context = json.load(file)

xlst = xlsxt.ExcelTemplateProcessor("./examples/demo.xlsx")
xlst.process_template(context)
xlst.save("./examples/output.xlsx")

print("****** Diffing with expected output_expected.xlsx (just value and formulas)")

diffs = xlsxt.compare_workbooks("./examples/output_expected.xlsx", "./examples/output.xlsx", compare_style=False)

for sheet_and_row, diff in diffs.items():
    print("sheet", sheet_and_row[0], "line:",sheet_and_row[1])
    print("\t", diff)

if diffs:
    raise Exception("some diff with expected rendering")