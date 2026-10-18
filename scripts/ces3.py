def bubble():
    # Bubble sort implementation
    arr = [64, 34, 25, 12, 22, 11, 90]
    n = len(arr)

    for i in range(n):
        # Last i elements are already sorted
        for j in range(0, n - i - 1):
            # Swap if the element found is greater than the next element
            if arr[j] > arr[j + 1]:
                arr[j], arr[j + 1] = arr[j + 1], arr[j]

    return arr


print("Sorted array is:", bubble())
