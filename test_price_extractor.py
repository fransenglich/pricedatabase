import unittest
import pickle

import price_extractor

# A set of real-scenario test cases.
#
# We pickle and check the DataFrames, but formatting of data
# in the outputted Excel file might be wrong, so manually
# checking it is also necessary.
class TestPriceExtractor(unittest.TestCase):

    test_cases = [
                  # Somewhat atomic.
                  ("header",        ["tests/in_header.xlsx"]),
                  ("two_rows",      ["tests/in_two_rows.xlsx"]),
                  ("three_rows",    ["tests/in_three_rows.xlsx"]),
                  ("space_entry",   ["tests/in_space_entry.xlsx"]),
                  ("no_quantity",   ["tests/in_no_quantity.xlsx"]),
                  ("simple",        ["tests/in_simple.xlsx"]),
                  ("simple2",       ["tests/in_simple2.xlsx"]),
                  ("simple3",       ["tests/in_simple3.xlsx"]),
                  ("simple12",      ["tests/in_simple.xlsx", "tests/in_simple2.xlsx"]),

                  # Complex.
                  ("real1",         ["tests/in_real1.xlsx"]),
                  ("hyllie",        ["tests/in_hyllie.xlsx"]),
                  ("complex1",      ["tests/in_complex1.xlsx"]),
                  ("actual_run",    price_extractor.input_files)
                  ]

    def test_compare(self):
        # Dangerous setting.
        update_all_baselines: bool = False

        for tc in self.test_cases:
            print(f"------------ TESTING {tc[0]} ------------")
            df = price_extractor.read_Xlsx_files(tc[1])

            out_filename = f'tests/baseline_{tc[0]}.pickle'

            def write_baseline(fn):
                with open(fn, "wb") as wf:
                    pickle.dump(df, wf)


            if update_all_baselines:
                write_baseline(out_filename)

            baseline = None

            # We can't compare the Excel files, but we write them out
            # to simplify debugging and diagnostics. Also good to at
            # least run the code path.
            price_extractor.to_excel(df, f'tests/out_{tc[0]}.xlsx')

            try:
                with open(out_filename, "rb") as f:
                    baseline = pickle.load(f)
            except:
                print(df.to_string())
                write_baseline("tests/NEW.pickle")
                print("No baseline found. Wrote out the data in NEW.pickle.")
                raise

            try:
                diff = baseline.compare(df)
            except:
                write_baseline("tests/ACTUAL.pickle")
                print("Actual in ACTUAL.pickle.")
                raise

            if not diff.empty:
                print(df.to_string())
                print(diff)
                write_baseline("tests/ACTUAL.pickle")
                print("Actual in ACTUAL.pickle.")

            self.assertTrue(diff.empty)

    def test_main(self):
        self.assertEqual(0, price_extractor.main())

if __name__ == "__main__":
    unittest.main()
