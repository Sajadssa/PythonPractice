
# -*- coding: utf-8 -*-
import sys
import os

# تنظیم encoding برای Windows
if os.name == 'nt':  # Windows
    import subprocess
    subprocess.run('chcp 65001', shell=True, capture_output=True)
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

fruits = {"apple", "banana", "cherry", "orange"}
other_fruits = {"موز", "سیب", "انگور", "گیلاس", "لیمو"}

union = fruits.union(other_fruits)
print("اتحاد:", union)

intersection = fruits.intersection(other_fruits)
print("اشتراک:", intersection)


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

# بهترین روش برای پیمایش روی دیکشنری# مجموعه با چند عدد
my_set = {1, 2, 3}
print(my_set)  # خروجی: {1, 2, 3}

# مجموعه از رشته‌ها# ایجاد مجموعه
fruits = {"سیب", "موز", "سیب", "پرتقال"}  # "سیب" تکراری حذف می‌شود
print(fruits)  # خروجی ممکن: {'پرتقال', 'سیب', 'موز'} (ترتیب نامشخص)

# اضافه کردن
fruits.add("کیوی")
print(fruits)  # خروجی ممکن: {'کیوی', 'پرتقال', 'سیب', 'موز'}

# حذف
fruits.remove("موز")
print(fruits)  # خروجی ممکن: {'کیوی', 'پرتقال', 'سیب'}

# عملیات مجموعه‌ای
other_fruits = {"انگور", "پرتقال", "هلو"}
union = fruits.union(other_fruits)
print(union)  # خروجی ممکن: {'کیوی', 'انگور', 'پرتقال', 'هلو', 'سیب'}

intersection = fruits.intersection(other_fruits)
print(intersection)  # خروجی ممکن: {'پرتقال'}
fruit_set = {"apple", "banana", "orange"}
print(fruit_set)  # خروجی: {'apple', 'banana', 'orange'}

# ساخت مجموعه از یک لیست (تبدیل لیست به مجموعه)
numbers = [1, 2, 2, 3, 4]
unique_numbers = set(numbers)
print(unique_numbers)  # خروجی: {1, 2, 3, 4}
print("\nپیمایش روی کلید و مقدار:")
for key, value in person.items():
    print(f"کلید: {key}, مقدار: {value}")
#  please explain about set in python?
# ایجاد مجموعه
fruits = {"سیب", "موز", "سیب", "پرتقال"}  # "سیب" تکراری حذف می‌شود
print(fruits)  # خروجی ممکن: {'پرتقال', 'سیب', 'موز'} (ترتیب نامشخص)

# اضافه کردن
fruits.add("کیوی")
print(fruits)  # خروجی ممکن: {'کیوی', 'پرتقال', 'سیب', 'موز'}

# حذف
fruits.remove("موز")
print(fruits)  # خروجی ممکن: {'کیوی', 'پرتقال', 'سیب'}

# عملیات مجموعه‌ای
other_fruits = {"انگور", "پرتقال", "هلو"}
union = fruits.union(other_fruits)
print(union)  # خروجی ممکن: {'کیوی', 'انگور', 'پرتقال', 'هلو', 'سیب'}

intersection = fruits.intersection(other_fruits)
print(intersection)  # خروجی ممکن: {'پرتقال'}

my_set = {1, 2, 3}
print(f"مجموعه اولیه: {my_set}")

# افزودن یک عنصر
my_set.add(4)
print(f"پس از افزودن 4: {my_set}")

# افزودن عنصری که از قبل وجود دارد (بدون تغییر)
my_set.add(2)
print(f"پس از افزودن 2 (موجود): {my_set}")

print("-" * 20)

# حذف یک عنصر با remove()
my_set.remove(3)
print(f"پس از حذف 3 با remove(): {my_set}")

# تلاش برای حذف عنصر ناموجود با remove() (اگر فعال شود خطا می‌دهد)
# try:
#     my_set.remove(10)
# except KeyError as e:
#     print(f"خطا هنگام حذف 10 با remove(): {e}")

print("-" * 20)

# حذف یک عنصر با discard()
my_set.discard(1)
print(f"پس از حذف 1 با discard(): {my_set}")

# تلاش برای حذف عنصر ناموجود با discard() (بدون خطا)
my_set.discard(10)
print(f"پس از حذف 10 با discard() (ناموجود): {my_set}")

print("-" * 20)

# حذف یک عنصر دلخواه با pop()
# توجه: خروجی pop می‌تواند در هر بار اجرا متفاوت باشد چون مجموعه‌ها ترتیب ندارند
if my_set: # مطمئن می‌شویم مجموعه خالی نیست قبل از pop
    popped_element = my_set.pop()
    print(f"پس از pop(): {my_set}, عنصر حذف شده: {popped_element}")
else:
    print("مجموعه خالی است، نمی‌توان pop کرد.")


print("-" * 20)

# خالی کردن مجموعه با clear()
my_set.clear()
print(f"پس از clear(): {my_set}")