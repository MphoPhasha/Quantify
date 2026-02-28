import unittest
from quantify import labelFormat, pipeTypeFormat, calculate_slopes, OutsideDiameter_Sewer

class TestQuantify(unittest.TestCase):
    
    def test_label_format(self):
        self.assertEqual(labelFormat(' "MH1" '), "MH1")
        self.assertEqual(labelFormat('MH2'), "MH2")
        self.assertEqual(labelFormat('  "  MH3  "  '), "MH3")

    def test_pipe_type_format(self):
        self.assertEqual(pipeTypeFormat(' "100mm uPVC" '), "100mm uPVC")
        self.assertEqual(pipeTypeFormat('200mm Concrete'), "200mm Concrete")

    def test_calculate_slopes(self):
        # 10m length, 1m fall = 0.1 slope
        invert_levels = [100.0, 99.0]
        chainages = [0.0, 10.0]
        file_numbers = ["001", "001"]
        slopes = calculate_slopes(invert_levels, chainages, file_numbers)
        self.assertEqual(len(slopes), 1)
        self.assertAlmostEqual(slopes[0], 0.1)
        
        # Zero length should result in 0 slope to avoid div by zero
        invert_levels = [100.0, 99.0]
        chainages = [10.0, 10.0]
        slopes = calculate_slopes(invert_levels, chainages, file_numbers)
        self.assertEqual(slopes[0], 0.0)

    def test_outside_diameter_sewer(self):
        pipe_types = [["110mm uPVC Class 34", "160mm uPVC Class 34"]]
        od = OutsideDiameter_Sewer(pipe_types)
        self.assertEqual(od, [[110.0, 160.0]])
        
        # Handle unexpected format gracefully
        pipe_types = [["Invalid"]]
        od = OutsideDiameter_Sewer(pipe_types)
        self.assertEqual(od, [[0.0]])

if __name__ == '__main__':
    unittest.main()
