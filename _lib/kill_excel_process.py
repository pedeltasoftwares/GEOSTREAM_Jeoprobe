"""
MATA TODOS LOS PROCESOS DE EXCEL
"""
import psutil

def kill_excel_processes():

    # Obtenemos todos los procesos para identificar los secundarios de Excel
    all_processes = {p.pid: p for p in psutil.process_iter(['pid', 'ppid', 'name'])}
    
    for process in all_processes.values():
        try:
            # Verificamos si el proceso principal es EXCEL.EXE
            if process.info['name'] and 'EXCEL.EXE' in process.info['name'].upper():
                print(f"Cerrando proceso principal: {process.info['name']} (PID: {process.pid})")
                process.terminate()
                # Identificamos y terminamos los procesos secundarios
                for child_pid, child_process in all_processes.items():
                    if child_process.info['ppid'] == process.pid:
                        print(f"Cerrando proceso secundario: {child_process.info['name']} (PID: {child_process.pid})")
                        child_process.terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass

