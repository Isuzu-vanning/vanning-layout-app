
from プロトタイプ import Container, PARTS_MASTER, Item

def test_counts():
    print("--- Testing Empty Container ---")
    c = Container("20ft")
    # Empty container
    counts = c.get_loadable_counts(PARTS_MASTER)
    print(f"Empty container suggestions: {counts}")

    print("\n--- Testing Container with 1 Item ---")
    c = Container("20ft")
    item = Item("BOX-L", PARTS_MASTER["BOX-L"], "test-1")
    c.load_items([item])
    counts = c.get_loadable_counts(PARTS_MASTER)
    print(f"Container with 1 Large Box suggestions: {counts}")

if __name__ == "__main__":
    test_counts()
