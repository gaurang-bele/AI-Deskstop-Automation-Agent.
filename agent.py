from llm_parser import parse_command
from executor import execute_action

def process_command(command):

    action = parse_command(command)

    result = execute_action(action)

    return result