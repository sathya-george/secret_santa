import pandas as pd
import random

def read_excel(file_path):
    df = pd.read_excel(file_path)
    employees = df.to_dict(orient='records')
    return employees

def assign_secret_santa(employees):
    names = [employee['Employee_Name'] for employee in employees]
    random.shuffle(names)
    
    valid_assignment = False
    
    while not valid_assignment:
        secret_childs = names[:]
        random.shuffle(secret_childs)
        
        valid_assignment = all(employee['Employee_Name'] != secret_child 
                               for employee, secret_child in zip(employees, secret_childs))
        

    for i, employee in enumerate(employees):
        employee['Secret_Child_Name'] = secret_childs[i]
        for emp in employees:
            if emp['Employee_Name'] == secret_childs[i]:
                employee['Secret_Child_EmailID'] = emp['Employee_EmailID']
                break
                
    return employees

def write_excel(employees, output_file_path):
    df = pd.DataFrame(employees)
    df.to_excel(output_file_path, index=False)

def main(input_excel, output_excel):
    employees = read_excel(input_excel)
    assigned_employees = assign_secret_santa(employees)
    write_excel(assigned_employees, output_excel)


input_excel = 'Employee_List.xlsx'
output_excel = 'Secret_Santa_Game_Result_2024.xlsx'


main(input_excel, output_excel)





