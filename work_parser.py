import openpyxl

class Calculator:
    def __init__(self, file_name_read):
        self.file_name_read = file_name_read

    def parse_excel(self):
        workbook = openpyxl.load_workbook(self.file_name_read)
        sheet = workbook.active

        user_input_pol = input("Please input POL: ").strip().lower()
        user_input_container = input("Please input container type (20DC/40HQ): ").strip().lower()
        user_input_dropoff = input("Please input drop off place: ").strip().lower()

        for row in sheet.iter_rows(min_row=2, values_only=True): 
            country, pol, pod, container, railway, freight, unloading, service, shipping, dropoff = row

            if pol.strip().lower() == user_input_pol and dropoff.strip().lower() == user_input_dropoff:
                if user_input_container == container:
                    if user_input_container == "20dc":
                        print(f"Match found for 20DC: {{'country': {country}, 'pol': {pol}, 'pod': {pod}, 'container': {container}, \n"
                              f"'railway rate': {railway}, 'freight price': {freight}, 'unloading price': {unloading}, \n"
                              f"'shipping line': {shipping}, 'dropoff place': {dropoff}}}")
                    elif user_input_container == "40hq":
                        print(f"Match found for 40HQ: {{'country': {country}, 'pol': {pol}, 'pod': {pod}, 'container': {container}, \n"
                              f"'railway rate': {railway}, 'freight price': {freight}, 'unloading price': {unloading}, \n"
                              f"'shipping line': {shipping}, 'dropoff place': {dropoff}}}")

calc = Calculator(file_name_read='calc_draft.xlsx') 
calc.parse_excel()