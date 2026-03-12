from tools.excel_tools import delete_rows

def execute_action(action):

    app = action["app"]
    act = action["action"]
    params = action["parameters"]

    if app == "excel" and act == "delete_rows":
        return delete_rows(params["rows"])

    return "Action not supported"