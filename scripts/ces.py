import pandas as pd

# beizhu ceshji


def add_data():
    # Create a sample DataFrame
    data = {"Name": ["Alice", "Bob", "Charli中国"], "Age": [25, 30, 35], "City": ["New York", "Los Angeles", "Chicago"]}
    df = pd.DataFrame(data)

    # Add a new column with default values
    df["Country"] = "USA"

    return df


if __name__ == "__main__":
    df = add_data()
    print(df)


ces = [x for x in range(5)]  # This line seems to be incomplete and may not serve a purpose
for i in ces:
    print(f"Iteration {i + 1}:")
    df = add_data()
    print(df)
    print()  # Print a newline for better readability

ls = [
    1,
    2,
    52,
    72,
    11,
    21,
]

print(list(range(5)))


def bubble_sort(arr):
    n = len(arr)
    for i in range(n):
        swapped = False
        for j in range(0, n - i - 1):
            if arr[j] > arr[j + 1]:
                arr[j], arr[j + 1] = arr[j + 1], arr[j]
                swapped = True
        if not swapped:  # 优化：已有序提前退出
            break
    return arr


print(bubble_sort(ls))


print("hello")


print("hello world")


print("hello world3333")


def bubble_sort(arr):
    n = len(arr)
    for i in range(n):
        swapped = False
        for j in range(0, n - i - 1):
            if arr[j] > arr[j + 1]:
                arr[j], arr[j + 1] = arr[j + 1], arr[j]
                swapped = True
        if not swapped:  # 优化：已有序提前退出
            print("Array is already sorted. Exiting early.")
            break
    return arr


print(f"Sorted array: {bubble_sort(ls)}")
