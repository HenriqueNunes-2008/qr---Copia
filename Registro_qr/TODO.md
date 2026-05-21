# TODO: Add Admin Area for Managing Areas, Projects, Employees, QR Codes, and Budgets

## Steps to Complete

- [x] Create config.json with initial areas and projetos
- [x] Create employees.json with initial employee list
- [x] Modify app.py to load config and employees from JSON files
- [x] Update /static/config.js route to return dynamic config
- [x] Add /api/employees route to return employee list
- [x] Create templates/admin.html with forms for add/delete areas, projects, and add employees
- [x] Add /admin route to render admin.html
- [x] Add POST routes for /admin/add_area, /admin/delete_area, /admin/add_projeto, /admin/delete_projeto, /admin/add_employee
- [x] Update templates/qrcodes.html to fetch employees dynamically and generate QR codes
- [x] Test the admin functionality
- [x] Add admin login button on main page with login/password prompt
- [x] Add /login route to verify credentials and redirect to /admin
- [x] Create orcamentos.json for budget data
- [x] Add functions to load/save orcamentos in app.py
- [x] Add /admin/add_orcamento route to set budgets for area/project/number
- [x] Add /api/orcamentos route to return budget data
- [x] Update admin.html to include budget definition form with dropdowns for areas/projects
- [x] Modify criar_planilha_se_nao_existir to create "Gráficos" sheet in Excel
- [x] Add atualizar_graficos function to calculate hours worked vs budgeted and create bar chart
- [x] Call atualizar_graficos when saving registrations
- [x] Format headers and column widths in "Gráficos" sheet like "Registros" sheet
