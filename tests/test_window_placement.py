import unittest

from window_placement import centered_bounds


class PopupPlacementTests(unittest.TestCase):
    def test_center_over_parent_instead_of_outside_left_edge(self):
        self.assertEqual(centered_bounds((0, 0, 1100, 800), (425, 255), (0, 0, 1920, 1040)),
                         (337, 272, 425, 255))

    def test_all_edges_and_taskbar_are_respected(self):
        for parent in ((-500, -400, 100, 100), (1800, 1000, 2400, 1500)):
            x, y, w, h = centered_bounds(parent, (425, 255), (0, 40, 1920, 1040))
            self.assertGreaterEqual(x, 0)
            self.assertGreaterEqual(y, 40)
            self.assertLessEqual(x + w, 1920)
            self.assertLessEqual(y + h, 1040)

    def test_negative_monitor_coordinates(self):
        self.assertEqual(centered_bounds((-1600, -900, -500, -100), (425, 255),
                                         (-1920, -1080, 0, 0)), (-1263, -628, 425, 255))

    def test_small_work_area_limits_size(self):
        self.assertEqual(centered_bounds((0, 0, 1100, 800), (425, 255), (0, 0, 320, 200)),
                         (0, 0, 320, 200))
