class Node:
    def __init__(self):
        self.data = "01/01/2024"
        self.value = 0
        self.is_negative = False
        self.operation = ""

    def get_data(self, line):

        self.data = line[0:10]

        index_value = line.find('R$')

        self.operation = line[10 : index_value]
        i = 0
        while (self.operation[i].isdecimal() == False):
            i += 1

        self.operation = self.operation[0:i]

        
        index_value += 3

        if line[index_value] == '-':
            self.is_negative = True
            index_value += 1
        
        line = line.replace(",", "*")
        line = line.replace(".", ",")
        line = line.replace("*", ".")

        number = ""

        while line[index_value] != "R":
            number = number + line[index_value]
            index_value += 1
        
        if len(number.strip()) > 6:
            number = str(number).replace(",", "")
        
        self.value = float(number)

    def cut_operation_code(self, key_operation):
        self.operation = key_operation

    
    def sumValue(self, value):
        self.value = float(self.value) + float(value)


    def node_to_string(self):
        return f"{self.data} {self.operation} R$ {self.value}"

    def show_node(self):
        print(f"Node - Data {self.data} - Value {self.value} {self.is_negative} - {self.operation}")
        pass

    def replace_comma_dot(self):
        self.value = str(self.value).replace(".", "*")
        self.value = str(self.value).replace(",", ".")
        self.value = str(self.value).replace("*", ",")
    