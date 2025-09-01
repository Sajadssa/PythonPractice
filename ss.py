print ("hello sajad");
name ="sajad"
print (name);a = 28  
b = 1.5  
c = "Hello!"  
d = True  
e = None  ; print(a,b,c,d);
name_Lname=input("Name/Lname :,");
print(name_Lname+name);
print(f" bold"+name);
a=int(input("number1:"));
b=int(input("number2:"));
if 1<=a<=20 and 1<=b<=20:
    print(a+b);
elif a % 2==0 and 20<=b<100:
    print(b-a);
else:
 print("you can not sum a,b -- does not subscribe in rang");
 

 


# In Python, a dictionary is a built-in data type that stores a collection of key-value pairs.
# It is also known as an associative array or a hash map in other programming languages.
# Dictionaries are unordered, mutable, and indexed by keys.

# Key characteristics of Python dictionaries:
# 1. Key-Value Pairs: Dictionaries store data as keys and their associated values.
# 2. Unordered: The items in a dictionary are not stored in any specific order.
# 3. Mutable: You can add, modify, or remove items from a dictionary.
# 4. Unique Keys: Each key in a dictionary must be unique.
# 5. Dynamic: Dictionaries can grow or shrink in size as needed.

# Creating a dictionary
# You can create a dictionary by enclosing a comma-separated list of key-value pairs in curly braces {}.
# A colon : is used to separate each key from its value.

# Empty dictionary
my_dict = {}
print(f"Empty dictionary: {my_dict}")

# Dictionary with integer keys
my_dict = {1: 'apple', 2: 'ball'}
print(f"Dictionary with integer keys: {my_dict}")

# Dictionary with mixed keys
my_dict = {'name': 'John', 1: [2, 4, 3]}
print(f"Dictionary with mixed keys: {my_dict}")

# Creating a dictionary using the dict() constructor
my_dict = dict({1:'apple', 2:'ball'})
print(f"Dictionary created with dict() constructor: {my_dict}")

my_dict = dict([(1,'apple'), (2,'ball')])
print(f"Dictionary created from a list of tuples: {my_dict}")


# Accessing elements from a dictionary
# You can access the value of a specific item by referring to its key name, in square brackets.
my_dict = {'name': 'John', 'age': 30, 'city': 'New York'}
print(f"Name: {my_dict['name']}")
print(f"Age: {my_dict['age']}")

# You can also use the get() method to access the value of a key.
# The get() method returns None if the key is not found, instead of raising a KeyError.
my_dict = {'name': 'John', 'age': 30, 'city': 'New York'}
print(f"Name: {my_dict.get('name')}")
print(f"Age: {my_dict.get('age')}")
print(f"Country: {my_dict.get('country')}")# Original problematic line (or similar):
# dict = {'some_key': 'some_value'} 

person = {
    "name": "سارا",
    "age": 28,
    "city": "تهران"
}

# گرفتن تمام کلیدها
all_keys = person.keys()
print(f"کلیدهای دیکشنری: {all_keys}")

# شما می‌توانید روی این کلیدها حلقه بزنید
print("\nپیمایش روی کلیدها:")
for key in person.keys():
    print(key)
person = {
    "name": "سارا",
    "age": 28,
    "city": "تهران"
}

# گرفتن تمام مقادیر
all_values = person.values()
print(f"مقادیر دیکشنری: {all_values}")

# پیمایش روی مقادیر
print("\nپیمایش روی مقادیر:")
for value in person.values():
    print(value)
person = {
    "name": "sara",
    "age": 28,
    "city": "iezh"
}

# گرفتن تمام آیتم‌ها
all_items = person.items()
print(f"آیتم‌های دیکشنری: {all_items}")

# بهترین روش برای پیمایش روی دیکشنری
print("\nپیمایش روی کلید و مقدار:")
for key, value in person.items():
    print(f"کلید: {key}, مقدار: {value}")
