"""Exercise the confidence image matching path used by DentWeb automation."""
import sys
import unittest


class ImageMatchingTests(unittest.TestCase):
    @unittest.skipUnless(sys.platform == "win32", "Windows build dependency lock")
    def test_confidence_finds_exact_image_without_desktop_input(self):
        import numpy as np
        import pyscreeze
        from PIL import Image

        random = np.random.default_rng(1729)
        pixels = random.integers(0, 256, size=(11, 13, 3), dtype=np.uint8)
        needle = Image.fromarray(pixels)
        screen = Image.new("RGB", (80, 70), (27, 31, 47))
        screen.paste(needle, (17, 23))
        match = pyscreeze.locate(needle, screen, confidence=0.99)
        self.assertEqual(tuple(match), (17, 23, 13, 11))


if __name__ == "__main__":
    unittest.main()
