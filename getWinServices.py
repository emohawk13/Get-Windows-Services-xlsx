import psutil
from openpyxl import Workbook

def get_running_services():
    running_services = []
    stopped_services = []
    for service in psutil.win_service_iter():
        service_info = {}
        service_info['Name'] = service.name()
        service_info['Status'] = service.status()
        service_info['Start Type'] = service.start_type()
        if service_info['Status'] == 'running':
            running_services.append(service_info)
        else:
            stopped_services.append(service_info)
    return running_services, stopped_services

def export_to_excel(running_data, stopped_data, filename):
    wb = Workbook()
    running_sheet = wb.active
    running_sheet.title = "Running Services"
    stopped_sheet = wb.create_sheet(title="Stopped Services")
    
    running_sheet.append(['Name', 'Status', 'Start Type'])
    for service_info in running_data:
        running_sheet.append([service_info['Name'], service_info['Status'], service_info['Start Type']])
    
    stopped_sheet.append(['Name', 'Status', 'Start Type'])
    for service_info in stopped_data:
        stopped_sheet.append([service_info['Name'], service_info['Status'], service_info['Start Type']])
    
    wb.save(filename)

running_services, stopped_services = get_running_services()
export_to_excel(running_services, stopped_services, 'services.xlsx')
print("Services data exported to 'services.xlsx' file.")
