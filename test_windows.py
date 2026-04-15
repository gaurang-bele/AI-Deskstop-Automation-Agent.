# Test script to debug window detection
# Run this while Chrome/Edge is open to see if it's detected

import pygetwindow as gw

print("=" * 50)
print("  WINDOW DETECTION TEST")
print("=" * 50)
print()

# Get all windows
all_windows = gw.getAllWindows()
print(f"Total windows found: {len(all_windows)}")
print()

# Show all windows with titles
print("All windows with titles:")
print("-" * 50)
for i, w in enumerate(all_windows):
    if w.title and len(w.title.strip()) > 0:
        print(f"  {i+1}. '{w.title}'")

print()
print("-" * 50)

# Check for Chrome
print("\nSearching for 'chrome'...")
chrome_windows = [w for w in all_windows if w.title and "chrome" in w.title.lower()]
if chrome_windows:
    print(f"  ✓ Found {len(chrome_windows)} Chrome window(s):")
    for w in chrome_windows:
        print(f"    - '{w.title}'")
else:
    print("  ✗ No Chrome windows found")

# Check for Edge
print("\nSearching for 'edge'...")
edge_windows = [w for w in all_windows if w.title and "edge" in w.title.lower()]
if edge_windows:
    print(f"  ✓ Found {len(edge_windows)} Edge window(s):")
    for w in edge_windows:
        print(f"    - '{w.title}'")
else:
    print("  ✗ No Edge windows found")

# Check for Google
print("\nSearching for 'google'...")
google_windows = [w for w in all_windows if w.title and "google" in w.title.lower()]
if google_windows:
    print(f"  ✓ Found {len(google_windows)} Google window(s):")
    for w in google_windows:
        print(f"    - '{w.title}'")
else:
    print("  ✗ No Google windows found")

print()
print("=" * 50)
print("TEST COMPLETE")
print("=" * 50)
