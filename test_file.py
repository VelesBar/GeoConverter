class controlled_execution:
     def __enter__(self):
        return self
     def __exit__(self, type, value, traceback):
        print('exit')

def aa():
    return 10

with controlled_execution() as thing:
    print(thing)
