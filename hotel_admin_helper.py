import datetime
# import openpyxl
import wx
import wx.adv
import os
from decimal import Decimal
from enum import Enum
from dataclasses import dataclass

app = wx.App()
curr_path = os.getcwd()

# interface

# constants
TODAY = datetime.date.today()
APP_EXIT = wx.NewIdRef()
APP_SAVE = wx.NewIdRef()
CHANGE_PRICES_FRAME = wx.NewIdRef()
TWOPLACES = Decimal(10) ** -2


# main frame


class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        super().__init__(parent, title=title)

        # Menu_bar
        menubar = wx.MenuBar()
        # File Menu
        fileMenu = wx.Menu()

        # save_menu = wx.MenuItem(fileMenu, APP_SAVE, "&Зберегти\tCtr+S", "Зберегти файл")
        # fileMenu.Append(save_menu)

        fileMenu.AppendSeparator()  # Separator

        exit_menu = wx.MenuItem(fileMenu, APP_EXIT, "Вихід\tCtrl+Q", "Вийти з додатку")
        fileMenu.Append(exit_menu)

        menubar.Append(fileMenu, "&Файл")

        # Settings menu
        settings_menu = wx.Menu()
        settings_price = wx.MenuItem(settings_menu, CHANGE_PRICES_FRAME, "&Ціни\tCtrl+P", "Встанови ціни "
                                                                                          "номери/тур.збір/Сніданок")
        settings_menu.Append(settings_price)

        menubar.Append(settings_menu, "&Налаштування")

        # menu binds
        self.SetMenuBar(menubar)
        self.Bind(wx.EVT_MENU, self.onSave, id=APP_SAVE)
        self.Bind(wx.EVT_MENU, self.onQuit, id=APP_EXIT)
        self.Bind(wx.EVT_MENU, self.show_settings_price_frame, id=CHANGE_PRICES_FRAME)

        # main elements
        panel = wx.Panel(self)

        main_sizer = wx.GridBagSizer(10, 10)

        button_confirm = wx.Button(panel, label="Створити підтвердження")
        button_bill = wx.Button(panel, label="Створити рахунок")
        button_bill_wc = wx.Button(panel, label="Створити рахунок безготівковий")
        button_act = wx.Button(panel, label="Створити акт")

        main_sizer.Add(button_confirm, pos=(0, 0), flag=wx.EXPAND | wx.LEFT | wx.TOP, border=2)
        main_sizer.Add(button_bill, pos=(0, 1), flag=wx.EXPAND | wx.LEFT | wx.TOP, border=2)
        main_sizer.Add(button_bill_wc, pos=(0, 2), flag=wx.EXPAND | wx.LEFT | wx.TOP, border=2)
        main_sizer.Add(button_act, pos=(0, 3), flag=wx.EXPAND | wx.LEFT | wx.TOP | wx.RIGHT, border=2)
        main_sizer.AddGrowableCol(0)
        main_sizer.AddGrowableCol(1)
        main_sizer.AddGrowableCol(2)
        main_sizer.AddGrowableCol(3)

        # binds
        button_confirm.Bind(wx.EVT_BUTTON, self.make_confirm, button_confirm)
        button_bill.Bind(wx.EVT_BUTTON, self.make_bill, button_bill)
        button_bill_wc.Bind(wx.EVT_BUTTON, self.make_bill_wc, button_bill_wc)
        button_act.Bind(wx.EVT_BUTTON, self.make_act, button_act)

        # mandatory

        guest_name_stat_txt = wx.StaticText(panel, label="Guest name:")  # str
        main_sizer.Add(guest_name_stat_txt, pos=(1, 0), flag=wx.LEFT, border=10)
        self.guest_name_text_ctrl = wx.TextCtrl(panel)
        main_sizer.Add(self.guest_name_text_ctrl, pos=(1, 1), flag=wx.EXPAND | wx.LEFT, border=10)
        # done

        date_make_stat_txt = wx.StaticText(panel, label="Make date:")  # datetime
        main_sizer.Add(date_make_stat_txt, pos=(2, 0), flag=wx.LEFT, border=10)
        self.date_make = wx.adv.DatePickerCtrl(panel, wx.ID_ANY, wx.DefaultDateTime,
                                               style=wx.adv.DP_DROPDOWN | wx.adv.DP_SHOWCENTURY)
        self.date_make.Bind(wx.adv.EVT_DATE_CHANGED, self.make_date_changed)
        main_sizer.Add(self.date_make, pos=(2, 1), flag=wx.EXPAND | wx.LEFT, border=10)
        # done

        checkin_date_stat_txt = wx.StaticText(panel, label="CheckIn date:")  # datetime
        main_sizer.Add(checkin_date_stat_txt, pos=(3, 0), flag=wx.LEFT, border=10)
        self.checkin_date = wx.adv.DatePickerCtrl(panel, wx.ID_ANY, wx.DefaultDateTime,
                                                  style=wx.adv.DP_DROPDOWN | wx.adv.DP_SHOWCENTURY)
        self.checkin_date.Bind(wx.adv.EVT_DATE_CHANGED, self.checkin_date_changed)
        main_sizer.Add(self.checkin_date, pos=(3, 1), flag=wx.EXPAND | wx.LEFT, border=10)
        # done

        checkout_date_stat_txt = wx.StaticText(panel, label="CheckOut date:")  # datetime
        main_sizer.Add(checkout_date_stat_txt, pos=(4, 0), flag=wx.LEFT, border=10)
        self.checkout_date = wx.adv.DatePickerCtrl(panel, wx.ID_ANY, wx.DefaultDateTime,
                                                   style=wx.adv.DP_DROPDOWN | wx.adv.DP_SHOWCENTURY)
        self.checkout_date.Bind(wx.adv.EVT_DATE_CHANGED, self.checkout_date_changed)
        main_sizer.Add(self.checkout_date, pos=(4, 1), flag=wx.EXPAND | wx.LEFT, border=10)
        self.tomorrow = TODAY + datetime.timedelta(days=1)
        self.checkout_date.SetValue(self.tomorrow)
        # done

        category_stat_txt = wx.StaticText(panel, label="Category:")  # combobox
        main_sizer.Add(category_stat_txt, pos=(5, 0), flag=wx.LEFT, border=10)
        # before use enum --> categories = ["Стандартний", "Класичний", "Напівлюкс", "Люкс", "ДеЛюкс"]
        categories = [room_category for room_category in RoomType]
        self.category = wx.ComboBox(panel, choices=categories, style=wx.CB_READONLY)
        main_sizer.Add(self.category, pos=(5, 1), flag=wx.EXPAND | wx.LEFT, border=10)
        self.category.Bind(wx.EVT_COMBOBOX, self.category_combobox)

        # done

        price_accomodation_PN_stat_txt = wx.StaticText(panel, label="Price per night:")  # float
        main_sizer.Add(price_accomodation_PN_stat_txt, pos=(6, 0), flag=wx.LEFT, border=10)
        self.price_accomodation_PN_text_ctrl = wx.TextCtrl(panel)
        main_sizer.Add(self.price_accomodation_PN_text_ctrl, pos=(6, 1), flag=wx.EXPAND | wx.LEFT, border=10)

        # done

        total_price_accomodation_stat_txt = wx.StaticText(panel, label="Total price accomodation:")  # float auto-score
        main_sizer.Add(total_price_accomodation_stat_txt, pos=(7, 0), flag=wx.LEFT, border=10)
        self.total_price_accomodation_text_ctrl = wx.TextCtrl(panel)
        main_sizer.Add(self.total_price_accomodation_text_ctrl, pos=(7, 1), flag=wx.EXPAND | wx.LEFT, border=10)

        count_of_guest_stat_txt = wx.StaticText(panel, label="Count of guest:")  # combobox
        main_sizer.Add(count_of_guest_stat_txt, pos=(8, 0), flag=wx.LEFT, border=10)
        count_of_guests = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"]
        self.count_of_guest = wx.ComboBox(panel, choices=count_of_guests, style=wx.CB_READONLY)
        main_sizer.Add(self.count_of_guest, pos=(8, 1), flag=wx.EXPAND | wx.LEFT, border=10)
        self.count_of_guest.SetValue("1")
        self.count_of_guest.Bind(wx.EVT_COMBOBOX, self.count_of_guest_combobox)
        # done

        admin_name_stat_txt = wx.StaticText(panel, label="Admin name:")  # combobox
        main_sizer.Add(admin_name_stat_txt, pos=(9, 0), flag=wx.LEFT, border=10)
        admins = ["Аліна", "Влад", "Сергій"]
        self.admin_name = wx.ComboBox(panel, choices=admins, style=wx.CB_READONLY)
        main_sizer.Add(self.admin_name, pos=(9, 1), flag=wx.EXPAND | wx.LEFT, border=10)
        self.admin_name.Bind(wx.EVT_COMBOBOX, self.admin_combobox)
        # done

        # optional

        tour_tax_stat_txt = wx.StaticText(panel, label="Tour tax total:")  # float auto-score
        main_sizer.Add(tour_tax_stat_txt, pos=(11, 0), flag=wx.LEFT, border=10)
        self.tour_tax_text_ctrl = wx.TextCtrl(panel)
        main_sizer.Add(self.tour_tax_text_ctrl, pos=(11, 1), flag=wx.EXPAND | wx.LEFT, border=10)
        self.tour_tax_text_ctrl.SetValue("33.50")
        self.tour_tax_checkbox = wx.CheckBox(panel)
        main_sizer.Add(self.tour_tax_checkbox, pos=(11, 2), flag=wx.ALL, border=5)
        self.tour_tax_checkbox.Bind(wx.EVT_CHECKBOX, self.checkbox_tour_tax, self.tour_tax_checkbox)
        self.tour_tax_checkbox.SetValue(True)

        count_of_rooms_stat_txt = wx.StaticText(panel, label="Count rooms:")  # combobox
        main_sizer.Add(count_of_rooms_stat_txt, pos=(12, 0), flag=wx.LEFT, border=10)
        self.count_of_rooms_list = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"]
        self.count_of_rooms = wx.ComboBox(panel, choices=self.count_of_rooms_list, style=wx.CB_READONLY)
        main_sizer.Add(self.count_of_rooms, pos=(12, 1), flag=wx.EXPAND | wx.LEFT, border=10)
        self.count_of_rooms.Bind(wx.EVT_COMBOBOX, self.count_of_rooms_combobox, self.count_of_rooms)
        self.count_of_rooms.SetValue("1")
        self.count_of_rooms.Disable()
        self.count_of_rooms_checkbox = wx.CheckBox(panel)
        main_sizer.Add(self.count_of_rooms_checkbox, pos=(12, 2), flag=wx.ALL, border=5)
        self.count_of_rooms_checkbox.Bind(wx.EVT_CHECKBOX, self.checkbox_count_of_rooms, self.count_of_rooms_checkbox)
        self.count_of_rooms_checkbox.SetValue(False)

        # Info bar:
        # info_panel = wx.Panel(panel)
        # guest_name_infopanel_stat_txt = wx.Tex

        # main_sizer.Add(info_panel)

        panel.SetSizer(main_sizer)

    # functions

    def make_date_changed(self, event):
        date_make = self.date_make.GetValue()
        date_make = date_make.Format("%d.%m.%y")
        print(f"make date is {date_make}")
        return date_make
    # done

    def checkin_date_changed(self, event):
        checkin_date = self.checkin_date.GetValue()
        self.get_duration_accomodation()
        checkin_date = checkin_date.Format("%d.%m.%y")
        print(f"Checkindate is {checkin_date}")
        self.total_price_accomodation()
        self.tour_tax_calculator()
        return checkin_date
    # done

    def checkout_date_changed(self, event):
        checkout_date = self.checkout_date.GetValue()
        self.get_duration_accomodation()
        checkout_date = checkout_date.Format("%d.%m.%y")
        print(f"Checkoutdate is {checkout_date}")
        self.total_price_accomodation()
        self.tour_tax_calculator()
        return checkout_date
    # done

    def get_duration_accomodation(self):
        date1 = self.checkin_date.GetValue()
        date2 = self.checkout_date.GetValue()
        delta = date2.Subtract(date1)
        duration_accomodation = delta.GetDays()
        return duration_accomodation
    # done

    def category_combobox(self, event):
        selected_category = self.category.GetValue()
        if selected_category:
            self.price_accomodation_PN_text_ctrl.SetValue(str(default_prices[selected_category]))
            self.total_price_accomodation()
        return selected_category
    # done

    def total_price_accomodation(self):
        duration_accomodation = self.get_duration_accomodation()
        price_accomodation_PN_text_ctrl = self.price_accomodation_PN_text_ctrl.GetValue()
        total_price_accomodation = int(duration_accomodation) * float(price_accomodation_PN_text_ctrl)
        total_price_accomodation = Decimal(total_price_accomodation).quantize(TWOPLACES)
        self.total_price_accomodation_text_ctrl.SetValue(str(total_price_accomodation))
        return total_price_accomodation

    def count_of_guest_combobox(self, event):
        count_of_guest = self.count_of_guest.GetValue()
        count_of_guest = int(count_of_guest)
        self.tour_tax_calculator()
        return count_of_guest

    def admin_combobox(self, event):
        selected_admin = self.admin_name.GetValue()
        return selected_admin

    def count_of_rooms_combobox(self, event):
        count_of_rooms = self.count_of_rooms.GetValue()
        return count_of_rooms

    # Checbox for tour tax and count of rooms
    def checkbox_tour_tax(self, event):
        tour_tax_checkbox = self.tour_tax_checkbox.GetValue()
        if tour_tax_checkbox:
            self.tour_tax_text_ctrl.Enable()
        else:
            self.tour_tax_text_ctrl.Disable()

    def tour_tax_calculator(self):
        tour_tax_total = Decimal(int(self.count_of_guest.GetValue()) * int(self.get_duration_accomodation()) *
                                 float(Prices.tourist_tax_price)).quantize(TWOPLACES)
        self.tour_tax_text_ctrl.SetValue(str(tour_tax_total))
        print(tour_tax_total, type(tour_tax_total))
        return tour_tax_total

    def checkbox_count_of_rooms(self, event):
        count_of_rooms_checbox = self.count_of_rooms_checkbox.GetValue()
        if count_of_rooms_checbox:
            self.count_of_rooms.Enable()
            self.count_of_rooms_combobox(event)
        else:
            self.count_of_rooms.SetValue(self.count_of_rooms_list[0])
            self.count_of_rooms_combobox(event)
            self.count_of_rooms.Disable()

    # makers

    def make_confirm(self, event):
        print("Gonna make confirm")

    def make_bill(self, event):
        print("Gonna make bill")

    def make_bill_wc(self, event):
        print("Gonna make bill without cash")

    def make_act(self, event):
        print("Gonna make act")

    # Setting price frame
    def show_settings_price_frame(self, event):
        setting_prie_frame = SettingPriceFrame(self, title="Налаштування цін за замовчанням")
        setting_prie_frame.SetSize(420, 310)
        setting_prie_frame.Centre()
        setting_prie_frame.Show()
        print("frame time")

    # Main menu bar func
    def onSave(self, event):
        print("gonna save")

    def onQuit(self, event):
        self.Close()


class SettingPriceFrame(wx.Frame):
    def __init__(self, parent, title):
        super().__init__(parent, title=title, style=wx.DEFAULT_FRAME_STYLE & ~wx.RESIZE_BORDER)

        spf_panel = wx.Panel(self)

        frame_sizer = wx.GridBagSizer(10, 10)

        top_text = wx.StaticText(spf_panel, label="Ціна за замовчанням")
        frame_sizer.Add(top_text, pos=(0, 1), flag=wx.EXPAND | wx.LEFT | wx.TOP | wx.RIGHT, border=10)

        standart_price_stat_txt = wx.StaticText(spf_panel, label="Стандарт ціна:")
        frame_sizer.Add(standart_price_stat_txt, pos=(1, 0), flag=wx.LEFT, border=10)
        self.standart_price_text_ctrl = wx.TextCtrl(spf_panel)
        self.standart_price_text_ctrl.SetValue(str(Prices.standart_price))
        frame_sizer.Add(self.standart_price_text_ctrl, pos=(1, 1), flag=wx.EXPAND | wx.LEFT, border=10)

        classic_price_stat_txt = wx.StaticText(spf_panel, label="Класичний ціна:")
        frame_sizer.Add(classic_price_stat_txt, pos=(2, 0), flag=wx.LEFT, border=10)
        self.classic_price_text_ctrl = wx.TextCtrl(spf_panel)
        self.classic_price_text_ctrl.SetValue(str(Prices.classic_price))
        frame_sizer.Add(self.classic_price_text_ctrl, pos=(2, 1), flag=wx.EXPAND | wx.LEFT, border=10)

        junior_suite_price_stat_txt = wx.StaticText(spf_panel, label="Напівлюкс ціна:")
        frame_sizer.Add(junior_suite_price_stat_txt, pos=(3, 0), flag=wx.LEFT, border=10)
        self.junior_suite_price_text_ctrl = wx.TextCtrl(spf_panel)
        self.junior_suite_price_text_ctrl.SetValue(str(Prices.junior_suite_price))
        frame_sizer.Add(self.junior_suite_price_text_ctrl, pos=(3, 1), flag=wx.EXPAND | wx.LEFT, border=10)

        suite_price_stat_txt = wx.StaticText(spf_panel, label="Люкс ціна:")
        frame_sizer.Add(suite_price_stat_txt, pos=(4, 0), flag=wx.LEFT, border=10)
        self.suite_price_text_ctrl = wx.TextCtrl(spf_panel)
        self.suite_price_text_ctrl.SetValue(str(Prices.suite_price))
        frame_sizer.Add(self.suite_price_text_ctrl, pos=(4, 1), flag=wx.EXPAND | wx.LEFT, border=10)

        delux_price_stat_txt = wx.StaticText(spf_panel, label="ДеЛюкс ціна:")
        frame_sizer.Add(delux_price_stat_txt, pos=(5, 0), flag=wx.LEFT, border=10)
        self.delux_price_text_ctrl = wx.TextCtrl(spf_panel)
        self.delux_price_text_ctrl.SetValue(str(Prices.delux_price))
        frame_sizer.Add(self.delux_price_text_ctrl, pos=(5, 1), flag=wx.EXPAND | wx.LEFT, border=10)

        tourist_tax_price_stat_txt = wx.StaticText(spf_panel, label="Туристичний збір:")
        frame_sizer.Add(tourist_tax_price_stat_txt, pos=(6, 0), flag=wx.LEFT, border=10)
        self.tourist_tax_price_text_ctrl = wx.TextCtrl(spf_panel)
        self.tourist_tax_price_text_ctrl.SetValue(str(Prices.tourist_tax_price))
        frame_sizer.Add(self.tourist_tax_price_text_ctrl, pos=(6, 1), flag=wx.EXPAND | wx.LEFT, border=10)

        breakfest_price_stat_txt = wx.StaticText(spf_panel, label="Сніданок ціна:")
        frame_sizer.Add(breakfest_price_stat_txt, pos=(7, 0), flag=wx.LEFT, border=10)
        self.breakfest_price_text_ctrl = wx.TextCtrl(spf_panel)
        self.breakfest_price_text_ctrl.SetValue(str(Prices.breakfest_price))
        frame_sizer.Add(self.breakfest_price_text_ctrl, pos=(7, 1), flag=wx.EXPAND | wx.LEFT, border=10)

        change_prices_button = wx.Button(spf_panel, label="Змінити\nціни\nза\nзамовчуванням")
        frame_sizer.Add(change_prices_button, pos=(0, 2), span=(9, 1), flag=wx.EXPAND | wx.ALL, border=10)
        change_prices_button.Bind(wx.EVT_BUTTON, self.change_prices, change_prices_button)

        spf_panel.SetSizer(frame_sizer)

    # functions
    def change_prices(self, event):
        default_prices[RoomType.Standart] = self.standart_price_text_ctrl.GetValue()
        default_prices[RoomType.Classic] = self.classic_price_text_ctrl.GetValue()
        default_prices[RoomType.JuniorSuite] = self.junior_suite_price_text_ctrl.GetValue()
        default_prices[RoomType.Suite] = self.suite_price_text_ctrl.GetValue()
        default_prices[RoomType.DeLux] = self.delux_price_text_ctrl.GetValue()
        Prices.tourist_tax_price = self.tourist_tax_price_text_ctrl.GetValue()
        Prices.breakfest_price = self.breakfest_price_text_ctrl.GetValue()


        print("Change prices")


# information
class RoomType(str, Enum):
    Standart = "Стандарт"
    Classic = "Класичний"
    JuniorSuite = "Напівлюкс"
    Suite = "Люкс"
    DeLux = "ДеЛюкс"
    TourTax = "Туристичний збір"
    Breakfest = "Сніданок"


default_prices: dict[RoomType, Decimal] = {
    RoomType.Standart: Decimal("1000.00"),
    RoomType.Classic: Decimal("1200.00"),
    RoomType.JuniorSuite: Decimal("1500.00"),
    RoomType.Suite: Decimal("1800.00"),
    RoomType.DeLux: Decimal("2200.00"),
    RoomType.TourTax: Decimal("33.50"),
    RoomType.Breakfest: Decimal("190.00")

}

@dataclass
class Client:
    name: str
    number_of_bill: int
    date_open: str
    date_make: str
    nights: int
    company_name: str


@dataclass
class Prices:
    standart_price = default_prices[RoomType.Standart]
    classic_price = default_prices[RoomType.Classic]
    junior_suite_price = default_prices[RoomType.JuniorSuite]
    suite_price = default_prices[RoomType.Suite]
    delux_price = default_prices[RoomType.DeLux]
    tourist_tax_price = Decimal("33.50")
    breakfest_price = Decimal("190.00")


# Main frame
main_frame = MyFrame(None, title="Admin_helper")
main_frame.SetSize(800, 700)
main_frame.Centre()
main_frame.Show()

app.MainLoop()
