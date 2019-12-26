import unittest
from unittest.mock import MagicMock, patch
import inv
import gui


class TestCase(unittest.TestCase):

    def test_select_dest_file(self):  # the argument is for testing purposes, avoiding dialog in test mode.
        """Selects file to save the current session operations. """
        file = r"C:\Users\jmishev\Desktop\inventory\REVIZIQ"
        gui.root.title(file)
        self.assertEqual(gui.root.title(), r"C:\Users\jmishev\Desktop\inventory\REVIZIQ")

    def setUp(self) -> None:
        inv.entries[0, 1].get = MagicMock()
        inv.entries[1, 1].get = MagicMock()
        inv.entries[2, 1].get = MagicMock()

    def tearDown(self) -> None:
        inv.entries[0, 1].get.dispose()
        inv.entries[1, 1].get.dispose()
        inv.entries[2, 1].get.dispose()

    def test_close_with_x(self):
        # mocking the entries
        original = inv.good_list.copy()
        with patch("inv.write_or_add"):
            with patch.dict(inv.good_list, {}):
                inv.entries[1, 1].get.return_value=None
                inv.entries[2, 1].get.return_value=None
                with self.assertRaises(SystemExit):
                    inv.close_with_x()

            assert inv.good_list == original

            inv.check_on_screen = MagicMock(return_value="Good to be entered")
            inv.entries[0, 1].get.return_value = "1"
            inv.entries[1, 1].get.return_value = "1"
            inv.entries[2, 1].get.return_value = "1"
            self.assertEqual(inv.close_with_x(), "Good entered")
            inv.check_on_screen = MagicMock(return_value=False)
            inv.good_list = {"1": 1}
            self.assertEqual(inv.close_with_x(), "Goods in the list entered no onscreen entered (close_with_x)")


    @patch("inv.write_to_excel")
    @patch("inv.add_to_excel")
    def test_write_or_add(self, mk1, mk2):
        gui.root.title = MagicMock(return_value="Програма за ревизия")
        self.assertEqual(inv.write_or_add(), "written to excel")
        gui.root.title.return_value="25.xlsX"
        self.assertEqual(inv.write_or_add(), "added to excel")

    @patch("inv.write_or_add")
    def test_close_with_button(self, m1):  # event is given by pushing the button
        """Closing with button "Запази  и затвори" """
        with patch("inv.exit"):
            with patch.dict(inv.good_list, {1:1}):
                inv.check_on_screen = MagicMock(return_value="No data")
                self.assertEqual(inv.close_with_button(), "Closing with button no onscreen")
                inv.check_on_screen.return_value = "Saving onscreen rejected"
                self.assertEqual(inv.close_with_button(), "Saving onscreen rejected")
                inv.check_on_screen.return_value = "Good to be entered"
                self.assertEqual(inv.close_with_button(), "onscreen entered")

    @patch("inv.enter_good")
    def test_check_on_screen(self, m):
        inv.messagebox.askyesno = MagicMock()
        inv.check_if_entries_are_numbers = MagicMock(return_value=True)
        inv.enter_good = MagicMock(return_value="Good entered")
        inv.entries[1, 1].get.return_value=None
        inv.entries[2, 1].get.return_value="2"
        self.assertEqual(inv.check_on_screen(), "No data")

        inv.entries[1, 1].get.return_value="3,44"
        inv.entries[2, 1].get.return_value="2"
        self.assertEqual(inv.check_on_screen(), "Good to be entered")


        inv.enter_good = MagicMock(return_value="Failed")
        self.assertRaises(ValueError)

        inv.messagebox.askyesno = MagicMock(return_value=False)
        self.assertEqual(inv.check_on_screen(), "Saving onscreen rejected")

    def test_check_if_entries_are_numbers(self):
        inv.entries[1, 1].get.return_value = "2"
        inv.entries[2, 1].get.return_value = "2"
        self.assertTrue(inv.check_if_entries_are_numbers())
        inv.entries[1, 1].get.return_value="m"
        inv.messagebox.showinfo = MagicMock(return_value=True)

    def test_highlight_wrong(self):
        inv.entries[0, 1].get.return_value = "m"
        inv.entries[1, 1].get.return_value = "m"
        inv.entries[2, 1].get.return_value="1"
        inv.highlight_wrong()
        self.assertEqual(inv.entries[0, 1]["highlightthickness"], 0)
        self.assertEqual(inv.entries[1, 1]["highlightthickness"], 2)
        self.assertEqual(inv.entries[2, 1]["highlightthickness"], 0)
        inv.entries[2, 1].return_value="1"
        inv.highlight_wrong()
        self.assertEqual(inv.entries[2, 1]["highlightthickness"], 0)

    @patch("inv.enter_good")
    def test_on_return(self, m):
        with patch("inv.root.focus_get"):
            gui.root.focus_get = MagicMock(return_value=inv.entries[0, 1])
            self.assertEqual(gui.on_return("Return"), inv.entries[1, 1])

            gui.root.focus_get.return_value = inv.entries[1, 1]
            self.assertEqual(gui.on_return("Return"), inv.entries[2, 1])

            gui.root.focus_get.return_value = inv.entries[2, 1]
            self.assertEqual(gui.on_return("Return"), inv.entries[0, 1])
            # TODO check if entry is created


    @patch("inv.check_if_entries_are_numbers")
    @patch("inv.create_good")
    @patch("inv.write_or_add")
    def test_enter_good(self, m, m1, m2):
        inv.messagebox.showinfo.return_value = True
        inv.entries[1, 1].get.return_value="3"
        inv.entries[2, 1].get.return_value="1"
        inv.check_if_entries_are_numbers.return_value=False
        self.assertEqual(inv.enter_good(), "Failed")
        inv.check_if_entries_are_numbers.return_value=True
        print(inv.check_if_entries_are_numbers())
        self.assertEqual(inv.enter_good(), "Good entered")

    def test_set_focus_on_grid_slave(self):
        self.assertEqual(gui.set_focus_on_grid_slave(1, 1), inv.entries[1, 1])
        self.assertEqual(gui.set_focus_on_grid_slave(2, 1), inv.entries[2, 1])


    def test_create_good(self):

        inv.entries[0, 1].get.return_value='a'
        inv.entries[1, 1].get.return_value="5"
        inv.entries[2, 1].get.return_value="6"
        inv.good_list_names = MagicMock()
        self.assertEqual(inv.create_good(), (" created ('a', 5.0, 6.0, 30.0)"))

        with patch.dict(inv.good_list, {'a': ('a', 5.0, 6.0, 30.0)}):
            inv.entries[0, 1].get.return_value='a'
            inv.entries[1, 1].get.return_value="5"
            inv.entries[2, 1].get.return_value="6"

            self.assertEqual(inv.create_good(), ("added ('a', 5.0, 6.0, 30.0)"))

            inv.entries[0, 1].get = MagicMock(return_value=None)
            inv.entries[1, 1].get = MagicMock(return_value="5")
            inv.entries[2, 1].get = MagicMock(return_value="6")
            self.assertEqual(inv.create_good(), (" No name created ('No Name good', 5.0, 6.0, 30.0)"))

    def test_add_quantity(self):
        inv.entries[0, 1].get.return_value='a'
        inv.entries[1, 1].get.return_value="6"
        inv.entries[2, 1].get.return_value="5"
        x = inv.Goods("m", 3.0, 5.0)
        y = inv.Goods("b", 3.0, 5.0)
        z = inv.Goods('a', 6.0, 5.0)
        with patch.dict(inv.good_list, {"m": x, "b": y, "a": z}):
            gui.v.get = MagicMock(return_value=1)
            self.assertEqual(inv.add_quantity(), "added ('a', 12.0, 5.0, 60.0)")
            inv.v.get = MagicMock(return_value=0)
            self.assertEqual(inv.add_quantity(), "added ('a', 6.0, 5.0, 30.0)")

    def test_calculate_totals(self):
        good = inv.Goods("No name", 6, 3)
        good1 = inv.Goods("No name1", 6, 5)
        with patch.dict(inv.good_list, {"No name": good, "No name1": good1}):
            self.assertEqual(inv.calculate_total(), 48)

    # def test_wrtie_to_database(self):
    #     with patch (inv.good_list, {"opa": ("opa", 3, 4, 12)}):
    #         inv.write_to_database()
    #         MySQL.my_cursor.execute("SELECT name FROM revisia WHERE quantity=3")
    #         my_result = MySQL.my_cursor.fetchall()
    #         self.assertEqual(my_result[0][0], "opa")


class TestCase2(unittest.TestCase):

    def test_arrow_moves(self):
        inv.event = MagicMock(keysym="Down", widget=inv.entries[1, 1])
        self.assertEqual(gui.arrow_moves(inv.event), (2, 1))

        inv.event = MagicMock(keysym="Up", widget=inv.entries[2, 1])
        self.assertEqual(gui.arrow_moves(inv.event), (1, 1))

        inv.event = MagicMock(keysym="Right", widget=inv.entries[1, 1])
        self.assertEqual(gui.arrow_moves(inv.event), (1, 2))

        inv.event = MagicMock(keysym="Up", widget=inv.entries[0, 1])
        self.assertEqual(gui.arrow_moves(inv.event), (0, 1))

        inv.event = MagicMock(keysym="Down", widget=inv.entries[2, 1])
        self.assertEqual(gui.arrow_moves(inv.event), (2, 1))