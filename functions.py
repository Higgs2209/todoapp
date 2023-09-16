FILEPATH = "todos.txt"


def get_todos(filepath=FILEPATH):
    """Read a Text Files and return the list of todos items"""
    with open(filepath, 'r') as file:
        todos_local = file.readlines()
    return todos_local


def write_todos(todos_arg, filepath=FILEPATH):
    """Write the to-do items list in the text file"""
    with open(filepath, 'w') as file:
        file.writelines(todos_arg)
