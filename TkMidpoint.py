#!/usr/bin/env python3
# -*- coding: utf-8 -*-

__version__ = "1.0.0"

import os
import re
import sys
import ssl
import platform
import subprocess
import webbrowser
import tkinter as tk
import urllib.request
import xml.etree.ElementTree as ET

from datetime import (datetime as dt, timedelta as td)
from tkinter.messagebox import showinfo
from tkinter.ttk import Treeview

if os.path.exists("records.csv"):
    pass
else:
    with open(
            file="records.csv",
            mode="w",
            encoding="utf-8"
    ) as csv:
        csv.write(
            "No,"
            "Name,"
            "Gender,"
            "Day,"
            "Month,"
            "Year,"
            "Hour,"
            "Minute,"
            "Latitude,"
            "Longitude\n"
        )


def select_module(
        name: str = "",
        file: list = [],
        path: str = ""
):
    if os.name == "posix":
        os.system(f"pip3 install {name}")
    elif os.name == "nt":
        if sys.version_info.minor == 6:
            if platform.architecture()[0] == "32bit":
                new_path = os.path.join(path, file[0])
                os.system(f"pip3 install {new_path}")
            elif platform.architecture()[0] == "64bit":
                new_path = os.path.join(path, file[1])
                os.system(f"pip3 install {new_path}")
        elif sys.version_info.minor == 7:
            if platform.architecture()[0] == "32bit":
                new_path = os.path.join(path, file[2])
                os.system(f"pip3 install {new_path}")
            elif platform.architecture()[0] == "64bit":
                new_path = os.path.join(path, file[3])
                os.system(f"pip3 install {new_path}")
        elif sys.version_info.minor == 8:
            if platform.architecture()[0] == "32bit":
                new_path = os.path.join(path, file[4])
                os.system(f"pip3 install {new_path}")
            elif platform.architecture()[0] == "64bit":
                new_path = os.path.join(path, file[5])
                os.system(f"pip3 install {new_path}")


try:
    from dateutil import tz
except ModuleNotFoundError:
    os.system("pip3 install python-dateutil")
    from dateutil import tz
try:
    from pytz import utc, timezone
except ModuleNotFoundError:
    os.system("pip3 install pytz")
    from pytz import utc, timezone
try:
    from timezonefinder import TimezoneFinder
except ModuleNotFoundError:
    os.system("pip3 install timezonefinder")
    from timezonefinder import TimezoneFinder
try:
    import xlsxwriter
    from xlsxwriter import workbook, worksheet
except ModuleNotFoundError:
    os.system("pip3 install XlsxWriter")
    import xlsxwriter
    from xlsxwriter import workbook, worksheet

PATH = os.path.join(os.getcwd(), "Eph", "Whl")

try:
    import swisseph as swe
except ModuleNotFoundError:
    select_module(
        name="pyswisseph",
        file=[i for i in os.listdir(PATH) if "pyswisseph" in i],
        path=PATH
    )
    import swisseph as swe

swe.set_ephe_path(os.path.join(os.getcwd(), "Eph"))

ALPHABETA = [chr(i) for i in range(65, 91)]

SIGNS = [
    "Aries",
    "Taurus",
    "Gemini",
    "Cancer",
    "Leo",
    "Virgo",
    "Libra",
    "Scorpio",
    "Sagittarius",
    "Capricorn",
    "Aquarius",
    "Pisces"
]

PLANETS = {
    "Sun": swe.SUN,
    "Moon": swe.MOON,
    "Mercury": swe.MERCURY,
    "Venus": swe.VENUS,
    "Mars": swe.MARS,
    "Jupiter": swe.JUPITER,
    "Saturn": swe.SATURN,
    "Uranus": swe.URANUS,
    "Neptune": swe.NEPTUNE,
    "Pluto": swe.PLUTO,
    "True": swe.TRUE_NODE,
    "Chiron": swe.CHIRON,
    "Juno": swe.JUNO,
    "Ceres": swe.CERES,
    "Pallas": swe.PALLAS,
    "Vesta": swe.VESTA
}

ASPECTS = {
    "Conjunction": {
        "degree": 0,
        "orb": "1\u00b0 0\' 0\"",
    },
    "Semi-Sextile": {
        "degree": 30,
        "orb": "1\u00b0 0\' 0\"",
    },
    "Semi-Square": {
        "degree": 45,
        "orb": "1\u00b0 0\' 0\"",
    },
    "Sextile": {
        "degree": 60,
        "orb": "1\u00b0 0\' 0\"",
    },
    "Quintile": {
        "degree": 72,
        "orb": "1\u00b0 0\' 0\"",
    },
    "Square": {
        "degree": 90,
        "orb": "1\u00b0 0\' 0\"",
    },
    "Trine": {
        "degree": 120,
        "orb": "1\u00b0 0\' 0\"",
    },
    "Sesquiquadrate": {
        "degree": 135,
        "orb": "1\u00b0 0\' 0\"",
    },
    "BiQuintile": {
        "degree": 144,
        "orb": "1\u00b0 0\' 0\"",
    },
    "Quincunx": {
        "degree": 150,
        "orb": "1\u00b0 0\' 0\"",
    },
    "Opposite": {
        "degree": 180,
        "orb": "1\u00b0 0\' 0\"",
    }
}

HOUSE_SYSTEMS = {
    "P": "Placidus",
    "K": "Koch",
    "O": "Porphyrius",
    "R": "Regiomontanus",
    "C": "Campanus",
    "E": "Equal",
    "W": "Whole Signs"
}

HSYS = "P"


def julday(
        year: int = 0,
        month: int = 0,
        day: int = 0,
        hour: int = 0,
        minute: int = 0,
        second: int = 0,
):
    if year < 0:
        t_frmt = "%d.%m.-%Y"
    else:
        t_frmt = "%d.%m.%Y"
    t_given = dt.strptime(f"{day}.{month}.{year}", t_frmt)
    t_limit = dt.strptime("15.10.1582", "%d.%m.%Y")
    if (t_limit - t_given).days > 0:
        calendar = swe.JUL_CAL
    else:
        calendar = swe.GREG_CAL
    jd = swe.julday(
        year,
        month,
        day,
        hour + (minute / 60) + (second / 3600),
        calendar
    )
    deltat = swe.deltat(jd)
    return {
        "JD": round(jd + deltat, 6),
        "TT": round(deltat * 86400, 1)
    }


def local_to_utc(
        year: int = 0,
        month: int = 0,
        day: int = 0,
        hour: int = 0,
        minute: int = 0,
        lat: float = 0,
        lon: float = 0
):
    local_zone = tz.gettz(
        str(timezone(TimezoneFinder().timezone_at(lat=lat, lng=lon)))
    )
    utc_zone = tz.gettz("UTC")
    global_time = dt.strptime(
        f"{year}-{month}-{day} {hour}:{minute}:00",
        "%Y-%m-%d %H:%M:%S"
    )
    local_time = global_time.replace(tzinfo=local_zone)
    utc_time = local_time.astimezone(utc_zone)
    return {
        "year": utc_time.year,
        "month": utc_time.month,
        "day": utc_time.day,
        "hour": utc_time.hour,
        "minute": utc_time.minute,
        "second": utc_time.second
    }


def progressed_time(
        nyear: int = 0,
        nmonth: int = 0,
        nday: int = 0,
        nhour: int = 0,
        nminute: int = 0,
        nsecond: int = 0,
        pyear: int = 0,
        pmonth: int = 0,
        pday: int = 0,
        phour: int = 0,
        pminute: int = 0,
        psecond: int = 0
):
    natal = dt.strptime(
        f"{nyear}.{nmonth}.{nday} {nhour}:{nminute}:{nsecond}",
        "%Y.%m.%d %H:%M:%S"
    )
    progressed = dt.strptime(
        f"{pyear}.{pmonth}.{pday} {phour}:{pminute}:{psecond}",
        "%Y.%m.%d %H:%M:%S"
    )
    time_delta = (progressed - natal).total_seconds() / 86400
    elapsed_day = round(time_delta / 365.24219893, 2)
    if elapsed_day > int(elapsed_day):
        elapsed_day = int(elapsed_day) + 1
    elapsed_time = (elapsed_day - time_delta / 365.24219893) * 86400
    return dt.strptime(
        f"{nyear}.{nmonth}.{nday}", "%Y.%m.%d"
    ) \
        + td(days=int(elapsed_day)) \
        + td(hours=natal.hour) \
        + td(minutes=natal.minute) \
        - td(minutes=int(elapsed_time / 60)) \
        - td(seconds=round(elapsed_time - int(elapsed_time / 60) * 60))


def convert_degree(degree: float = 0):
    for i in range(12):
        if i * 30 <= degree < (i + 1) * 30:
            return degree - (30 * i), [*SIGNS][i]


def reverse_convert_degree(degree: float = 0, sign: str = ""):
    return degree + 30 * [*SIGNS].index(sign)


def dd_to_dms(dd):
    degree = int(dd)
    minute = int((dd - degree) * 60)
    second = round(float((dd - degree - minute / 60) * 3600))
    return f"{degree}\u00b0 {minute}\' {second}\""


def dms_to_dd(dms):
    dms = dms.replace(" ", "")
    dms = dms.replace("\u00b0", " ").replace("\'", " ").replace("\"", " ")
    degree = int(dms.split(" ")[0])
    minute = float(dms.split(" ")[1]) / 60
    second = float(dms.split(" ")[2]) / 3600
    return degree + minute + second


class Chart:

    def __init__(
            self,
            jd: float = 0,
            lat: float = 0,
            lon: float = 0,
            hsys: str = ""
    ):
        self.jd = jd
        self.lat = lat
        self.lon = lon
        self.hsys = hsys

    def planet_pos(self, planet: int = 0):
        calc = convert_degree(
            degree=swe.calc_ut(self.jd, planet)[0]
        )
        return calc[1], reverse_convert_degree(calc[0], calc[1])

    def house_pos(self):
        house = []
        asc = 0
        degree = []
        for i, j in enumerate(swe.houses(
                self.jd, self.lat, self.lon,
                bytes(self.hsys.encode("utf-8")))[0]):
            if i == 0:
                asc += j
            degree.append(j)
            house.append((
                f"{i + 1}",
                j,
                f"{convert_degree(j)[1]}"))
        return house, asc, degree

    def patterns(self):
        planet_positions = []
        house_positions = []
        for i in range(12):
            house = [
                int(self.house_pos()[0][i][0]),
                self.house_pos()[0][i][-1],
                float(self.house_pos()[0][i][1]),
            ]
            house_positions.append(house)
        hp = [j[-1] for j in house_positions]
        for key, value in PLANETS.items():
            planet = self.planet_pos(planet=value)
            house = 0
            for i in range(12):
                if i != 11:
                    if hp[i] < planet[1] < hp[i + 1]:
                        house = i + 1
                        break
                    elif hp[i] < planet[1] > hp[i + 1] \
                            and hp[i] - hp[i + 1] > 240:
                        house = i + 1
                        break
                    elif hp[i] > planet[1] < hp[i + 1] \
                            and hp[i] - hp[i + 1] > 240:
                        house = i + 1
                        break
                else:
                    if hp[i] < planet[1] < hp[0]:
                        house = i + 1
                        break
                    elif hp[i] < planet[1] > hp[0] \
                            and hp[i] - hp[0] > 240:
                        house = i + 1
                        break
                    elif hp[i] > planet[1] < hp[0] \
                            and hp[i] - hp[0] > 240:
                        house = i + 1
                        break
            planet_info = [
                key,
                planet[0],
                planet[1],
                f"H{house}"
            ]
            planet_positions.append(planet_info)
        asc = house_positions[0] + ["H1"]
        asc[0] = "Asc"
        mc = house_positions[9] + ["H10"]
        mc[0] = "MC"
        planet_positions.extend([asc, mc])
        return planet_positions


class Record:

    def __init__(
            self,
            year: int = 0,
            month: int = 0,
            day: int = 0,
            hour: int = 0,
            minute: int = 0,
            lat: float = .0,
            lon: float = .0
    ):
        self.YEAR = year
        self.MONTH = month
        self.DAY = day
        self.LOCAL_HOUR = hour
        self.LOCAL_MINUTE = minute
        self.LAT = lat
        self.LON = lon
        self.UTC_YEAR = self.__utc()["year"]
        self.UTC_MONTH = self.__utc()["month"]
        self.UTC_DAY = self.__utc()["day"]
        self.UTC_HOUR = self.__utc()["hour"]
        self.UTC_MINUTE = self.__utc()["minute"]

    def __utc(self):
        return local_to_utc(
            year=self.YEAR,
            month=self.MONTH,
            day=self.DAY,
            hour=self.LOCAL_HOUR,
            minute=self.LOCAL_MINUTE,
            lat=self.LAT,
            lon=self.LON
        )

    def sidereal_time(self):
        jd = julday(
            year=self.UTC_YEAR,
            month=self.UTC_MONTH,
            day=self.UTC_DAY,
            hour=0,
            minute=0,
        )
        stamig = dt.strptime(
            dd_to_dms(swe.sidtime(jd["JD"])), '%H\u00b0 %M\' %S"'
        ) - td(seconds=jd["TT"])
        cgl_H = int(self.LON * 4 / 60)
        cgl_M = int(self.LON * 4 - (cgl_H * 60))
        cgl_S = int((self.LON * 4 - (cgl_H * 60) - int(cgl_M)) * 60)
        stc = int((self.UTC_HOUR / 24 + self.UTC_MINUTE / 1440) * 236)
        stc_M = int(stc / 60)
        stc_S = stc - (stc_M * 60)
        return td(hours=stamig.hour) \
            + td(minutes=stamig.minute) \
            + td(seconds=stamig.second) \
            + td(hours=self.UTC_HOUR) \
            + td(minutes=self.UTC_MINUTE) \
            + td(hours=cgl_H) \
            + td(minutes=cgl_M) \
            + td(seconds=cgl_S) \
            + td(minutes=stc_M) \
            + td(seconds=stc_S)

    def __progressed_time(
            self,
            year: int = 0,
            month: int = 0,
            day: int = 0,
    ):
        return progressed_time(
            nyear=self.UTC_YEAR,
            nmonth=self.UTC_MONTH,
            nday=self.UTC_DAY,
            nhour=self.UTC_HOUR,
            nminute=self.UTC_MINUTE,
            nsecond=0,
            pyear=year,
            pmonth=month,
            pday=day,
            phour=0,
            pminute=0,
            psecond=0
        )

    def natal(self, hsys: str = "P"):
        return Chart(
            lat=self.LAT,
            lon=self.LON,
            jd=julday(
                year=self.UTC_YEAR,
                month=self.UTC_MONTH,
                day=self.UTC_DAY,
                hour=self.UTC_HOUR,
                minute=self.UTC_MINUTE,
            )["JD"],
            hsys=hsys
        ).patterns()

    def progressed(
            self,
            year: int = 0,
            month: int = 0,
            day: int = 0,
            hsys: str = "P"
    ):
        ptime = self.__progressed_time(year=year, month=month, day=day)
        return Chart(
            lat=self.LAT,
            lon=self.LON,
            jd=julday(
                year=ptime.year,
                month=ptime.month,
                day=ptime.day,
                hour=self.UTC_HOUR,
                minute=self.UTC_MINUTE,
            )["JD"],
            hsys=hsys
        ).patterns()

    def transit(
            self,
            year: int = 0,
            month: int = 0,
            day: int = 0,
            hsys: str = "P"
    ):
        return Chart(
            lat=self.LAT,
            lon=self.LON,
            jd=julday(
                year=year,
                month=month,
                day=day,
                hour=0,
                minute=0,
            )["JD"],
            hsys=hsys
        ).patterns()


class App(tk.Frame):

    def __init__(self, master=None):
        super().__init__(master=master)
        self.pack(expand=True, fill=tk.BOTH)
        self.menu = tk.Menu(master=master)
        master.configure(menu=self.menu)
        self.record = tk.Menu(master=self, tearoff=False)
        self.interpretations = tk.Menu(master=self, tearoff=False)
        self.settings = tk.Menu(master=self, tearoff=False)
        self.help = tk.Menu(master=self, tearoff=False)
        self.menu.add_cascade(label="Record", menu=self.record)
        self.menu.add_cascade(
            label="Interpretations",
            menu=self.interpretations
        )
        self.menu.add_cascade(label="Settings", menu=self.settings)
        self.menu.add_cascade(label="Help", menu=self.help)
        self.record.add_command(
            label="Create",
            command=self.create_record
        )
        self.record.add_command(
            label="Open",
            command=self.open_record
        )
        self.interpretations.add_command(
            label="View",
            command=self.view_interpretations
        )
        self.settings.add_command(
            label="Orb Factor",
            command=self.choose_orb_factor
        )
        self.help.add_command(
            label="About",
            command=self.about
        )
        self.help.add_command(
            label="Check for updates",
            command=self.update_script
        )
        self.right_click = None
        self.treeview = None

    def popdown(self):
        if self.right_click:
            self.right_click.destroy()
            self.right_click = None

    def popup(
            self,
            event: tk.Event = None,
            treeview: Treeview = None,
            func: callable = None
    ):
        self.popdown()
        if treeview and not treeview.selection():
            return
        self.right_click = tk.Menu(master=None, tearoff=False)
        func()
        self.right_click.post(event.x_root, event.y_root)

    @staticmethod
    def max_char(event: None = None, limit: int = 0):
        if len(event.widget.get()) > limit:
            event.widget.delete(str(limit))

    @staticmethod
    def save_button_command(
            master: tk.Toplevel = None,
            entries: dict = {},
            create: bool = False,
            edit: list = [],
            treeview: Treeview = None,
            selected: tuple = ()
    ):
        error = False
        for k, v in entries.items():
            if k in ["Day", "Month", "Year", "Hour", "Minute"]:
                try:
                    int(v.get())
                except ValueError:
                    error = True
            elif k in ["Latitude", "Longitude"]:
                try:
                    float(v.get())
                except ValueError:
                    error = True
            elif k == "Name":
                if not v.get() or "," in v.get():
                    error = True
            elif k == "Gender":
                if not v.get():
                    error = True
        if error:
            showinfo(
                title="Warning",
                message="Fill all the entries with the "
                        "appropriate values."
            )
        else:
            try:
                dt.strptime(
                    f"{entries['Year'].get()}."
                    f"{entries['Month'].get()}."
                    f"{entries['Day'].get()} "
                    f"{entries['Hour'].get()}:"
                    f"{entries['Minute'].get()}:00",
                    f"%Y.%m.%d %H:%M:%S"
                )
            except ValueError:
                error = True
                showinfo(title="Warning", message="Invalid date.")

            try:
                TimezoneFinder().timezone_at(
                    lat=float(entries["Latitude"].get()),
                    lng=float(entries["Longitude"].get())
                )
            except:
                error = True
                showinfo(
                    title="Warning",
                    message="Invalid Latitude and Longitude."
                )
        if not error:
            row_data = f'"{entries["Name"].get()}",' \
                       f'"{entries["Gender"].get()}",' \
                       f'{int(entries["Day"].get())},' \
                       f'{int(entries["Month"].get())},' \
                       f'{int(entries["Year"].get())},' \
                       f'{int(entries["Hour"].get())},' \
                       f'{int(entries["Minute"].get())},' \
                       f'{float(entries["Latitude"].get())},' \
                       f'{float(entries["Longitude"].get())}\n'
            if create:
                with open("records.csv", "r") as f:
                    length = len(f.readlines())
                    row_data = f"{length}," + row_data
                with open("records.csv", "a+") as f:
                    f.write(row_data)
                    f.flush()
                    try:
                        if treeview:
                            treeview.insert(
                                parent="",
                                index=length,
                                values=[
                                    col.replace("\"", "") for col in
                                    row_data.replace("\n", "").split(",")
                                ]
                            )
                    except tk.TclError:
                        pass

            else:
                if treeview:
                    index = [
                        *treeview.get_children()
                    ].index(selected[0])
                    row_data = f"{treeview.item(selected)['values'][0]}," \
                               + row_data
                    treeview.delete(selected)
                    treeview.insert(
                        parent="",
                        index=index,
                        values=[
                            col.replace("\"", "")
                            for col in row_data.replace("\n", "").split(",")
                        ]
                    )
                    treeview.selection_set(
                        [*treeview.get_children()][index]
                    )
                edit = [
                    edit[0] + ",\"" + edit[1] + "\"",
                    "\"" + edit[2] + "\"",
                    *edit[3:-1], edit[-1] + "\n"
                ]
                if edit[-1].count("\n") > 1:
                    edit[-1] = edit[-1][:-1]
                with open("records.csv", "r") as f:
                    records = [i for i in f.readlines()]
                    index = records.index(",".join(edit))
                    records[index] = row_data
                with open("records.csv", "w") as f:
                    for i in records:
                        f.write(i)
                        f.flush()
            master.destroy()

    def record_panel(
            self,
            create: bool = False,
            edit: list = [],
            treeview: Treeview = None,
            selected: tuple = ()
    ):
        master = tk.Toplevel(master=None)
        master.geometry("300x300")
        master.resizable(width=False, height=False)
        if create:
            master.title("Create a record")
        if edit:
            master.title("Edit a record")
        entry_names = [
            "Name",
            "Gender",
            ["Day", "Month", "Year"],
            ["Hour", "Minute"],
            ["Latitude", "Longitude"]
        ]
        entries = {}
        var = tk.StringVar()
        for i, j in enumerate(entry_names):
            frame = tk.Frame(master=master)
            frame.pack()
            if isinstance(j, list):
                for k, m in enumerate(j):
                    sub_frame = tk.Frame(master=frame)
                    sub_frame.pack(side="left")
                    label = tk.Label(master=sub_frame, text=m)
                    label.pack()
                    if m in ["Month", "Day", "Hour", "Minute"]:
                        entry = tk.Entry(master=sub_frame, width=2)
                        entry.bind(
                            sequence="<KeyRelease>",
                            func=lambda event: self.max_char(
                                event=event,
                                limit=2
                            )
                        )
                    elif m == "Year":
                        entry = tk.Entry(master=sub_frame, width=4)
                        entry.bind(
                            sequence="<KeyRelease>",
                            func=lambda event: self.max_char(
                                event=event,
                                limit=4
                            )
                        )
                    else:
                        entry = tk.Entry(master=sub_frame, width=10)
                        entry.bind(
                            sequence="<KeyRelease>",
                            func=lambda event: self.max_char(
                                event=event,
                                limit=10
                            )
                        )
                    entry.pack()
                    entries[m] = entry
            else:
                label = tk.Label(master=frame, text=j)
                label.pack()
                if i == 0:
                    entry = tk.Entry(frame)
                    entry.pack()
                    entries[j] = entry
                elif i == 1:
                    option = tk.OptionMenu(frame, var, "M", "F", "N/A")
                    option.pack()
                    entries[j] = var
                elif i == 4:
                    entry = tk.Entry(frame)
                    entry.pack()
                    entries[j] = entry
        if edit:
            for i, j in zip(entries, edit[1:]):
                if i == "Gender":
                    entries[i].set(edit[2])
                else:
                    entries[i].insert(0, j)
        save_button = tk.Button(
            master=master,
            text="Save",
            command=lambda: self.save_button_command(
                master=master,
                entries=entries,
                create=create,
                edit=edit,
                treeview=treeview,
                selected=selected
            )
        )
        save_button.pack(pady=20)

    def create_record(self):
        self.record_panel(create=True, treeview=self.treeview)

    def sort_column(
            self,
            treeview: Treeview = None,
            col: int = 0,
            reverse: bool = False
    ):
        l = [
            (treeview.set(k, col), k)
            for k in treeview.get_children("")
        ]
        try:
            l.sort(key=lambda t: int(t[0]), reverse=reverse)
        except ValueError:
            l.sort(reverse=reverse)
        for index, (val, k) in enumerate(l):
            treeview.move(k, "", index)
        treeview.heading(
            column=col,
            command=lambda: self.sort_column(
                treeview=treeview,
                col=col,
                reverse=not reverse
            )
        )

    def heading(
            self,
            treeview: Treeview = None,
            col: int = 0,
            text: str = ""
    ):
        treeview.heading(
            column=f"#{col + 1}",
            text=text,
            command=lambda: self.sort_column(
                treeview=treeview,
                col=col,
                reverse=False
            )
        )

    def open_record(self):
        if os.path.exists("records.csv"):
            records = [
                i[:-1] for i in open("records.csv").readlines()[1:]
            ]
            master = tk.Toplevel(master=None)
            master.title("Open a record")
            master.resizable(width=False, height=False)
            frame = tk.Frame(master=master)
            frame.pack(expand=True, fill=tk.BOTH)
            y_scrollbar = tk.Scrollbar(master=frame, orient="vertical")
            y_scrollbar.pack(side="right", fill="y")
            columns = [
                "No",
                "Name",
                "Gender",
                "Day",
                "Month",
                "Year",
                "Hour",
                "Minute",
                "Latitude",
                "Longitude"
            ]
            search_entry = tk.Entry(
                master=frame,
                font="Default 10 italic",
            )

            search_entry.insert(
                index="0",
                string="Search a name..."
            )
            search_entry.pack()
            self.treeview = Treeview(
                master=frame,
                show="headings",
                columns=[f"#{i + 1}" for i in range(len(columns))],
                height=20,
                selectmode="extended",
                yscrollcommand=y_scrollbar.set
            )
            self.treeview.pack()
            y_scrollbar.configure(command=self.treeview.yview)
            search_entry.bind(
                sequence="<Button-1>",
                func=lambda event: self.delete_default_text(
                    event=event,
                    text="Search a name..."
                )
            )
            search_entry.bind(
                sequence="<KeyRelease>",
                func=lambda event: self.search_name(
                    event=event,
                    treeview=self.treeview
                )
            )
            search_entry.bind(
                sequence="<Button-3>",
                func=lambda event: self.popup(
                    event=event,
                    func=self.button_3_on_entry,
                    treeview=None
                )
            )
            search_entry.bind(
                sequence="<Control-KeyRelease-a>",
                func=lambda event: event.widget.select_range(0, tk.END)
            )
            frame.bind(
                sequence="<Button-1>",
                func=lambda event: self.insert_default_text(
                    entry=search_entry,
                    text="Search a name..."
                )
            )
            frame.bind(
                sequence="<Button-3>",
                func=lambda event: self.popdown()
            )
            self.treeview.bind(
                sequence="<Button-1>",
                func=lambda event: self.insert_default_text(
                    entry=search_entry,
                    text="Search a name..."
                )
            )
            self.treeview.bind(
                sequence="<Button-3>",
                func=lambda event: self.popup(
                    event=event,
                    func=lambda: self.button_3_on_treeview(self.treeview),
                    treeview=self.treeview
                )
            )
            for i, j in enumerate(records):
                self.treeview.insert(
                    parent="",
                    index=i,
                    values=[
                        col.replace("\"", "") for col in j.split(",")
                    ]
                )
            for i, j in enumerate(columns):
                self.treeview.column(
                    column=f"#{i + 1}",
                    minwidth=100,
                    width=100,
                    anchor=tk.CENTER
                )
                self.heading(
                    treeview=self.treeview,
                    col=i,
                    text=j
                )

    def delete_default_text(self, event: None = None, text: str = ""):
        self.popdown()
        if text in event.widget.get():
            event.widget.configure(font="Default 10")
            event.widget.delete("0", "end")

    def insert_default_text(self, entry: tk.Entry = None, text: str = ""):
        self.popdown()
        if not entry.get():
            entry.configure(font="Default 10 italic")
            entry.insert(
                index="0",
                string=text
            )

    @staticmethod
    def search_name(
            treeview: Treeview = None,
            event: tk.Event = None
    ):
        for i in treeview.get_children():
            if event.widget.get() == treeview.item(i)["values"][1]:
                treeview.selection_set(i)

    @staticmethod
    def delete_record(treeview: Treeview = None):
        selected = treeview.selection()
        if not selected:
            pass
        else:
            try:
                item = treeview.item(selected)["values"]
                item = [str(i) for i in item]
                item = [
                    item[0], "\"" + item[1] + "\"",
                             "\"" + item[2] + "\"",
                    *item[3:-1], item[-1] + "\n"
                ]
                if item[-1].count("\n") > 1:
                    item[-1] = item[-1][:-1]
                with open("records.csv", "r") as f:
                    new = [
                        i for i in f.readlines()
                        if i != ",".join(item)
                    ]
                    new = [new[0]] + [
                        ",".join([str(i + 1)] + j.split(",")[1:])
                        for i, j in enumerate(new[1:])
                    ]
                with open("records.csv", "w") as f:
                    for i in new:
                        f.write(i)
                        f.flush()
                for i, j in enumerate(treeview.get_children()):
                    if i >= int(item[0]) - 1:
                        treeview.delete(j)
                for i, j in enumerate(new[int(item[0]):]):
                    treeview.insert(
                        parent="",
                        index=int(item[0]) + i,
                        values=[
                            k.replace("\"", "")
                            for k in j.replace("\n", "").split(",")
                        ]

                    )
            except tk.TclError:
                pass

    def edit_record(self, treeview: Treeview = None):
        selected = treeview.selection()
        if not selected:
            pass
        else:
            try:
                edit = [str(i) for i in treeview.item(selected)["values"]]
                self.record_panel(
                    edit=edit,
                    treeview=treeview,
                    selected=selected
                )
            except tk.TclError:
                pass

    def button_3_on_treeview(
            self,
            treeview: Treeview = None
    ):
        midpoint = tk.Menu(master=self.right_click, tearoff=False)
        midpoint_list = tk.Menu(master=midpoint, tearoff=False)
        midpoint_aspects = tk.Menu(master=midpoint, tearoff=False)
        self.right_click.add_command(
            label="Delete",
            command=lambda: self.delete_record(treeview)
        )
        self.right_click.add_command(
            label="Edit",
            command=lambda: self.edit_record(treeview)
        )
        self.right_click.add_cascade(
            label="Export",
            menu=midpoint
        )
        midpoint.add_cascade(
            label="Midpoints",
            menu=midpoint_list
        )
        midpoint.add_cascade(
            label="Midpoint Aspects",
            menu=midpoint_aspects
        )
        midpoint.add_command(
            label="Midpoint Synastry Aspects",
            command=lambda: self.midpoint_synastry(
                treeview=treeview
            )
        )
        heading = lambda *args: \
            f"Aspects between midpoints of "\
            f"{args[0]} planets and {args[1]} planets"
        midpoint_list.add_command(
            label="Natal Midpoints",
            command=lambda: self.export_midpoints(
                treeview=treeview,
                heading="Natal Midpoints",
                row=7
            )
        )
        midpoint_list.add_command(
            label="Progressed Midpoints",
            command=lambda: self.export_midpoints(
                treeview=treeview,
                heading="Progressed Midpoints",
                row=8
            )
        )
        midpoint_aspects.add_command(
            label=heading("natal", "natal"),
            command=lambda: self.export_midpoint_aspects(
                treeview=treeview,
                heading=heading("natal", "natal")
            )
        )
        midpoint_aspects.add_command(
            label=heading("natal", "transit"),
            command=lambda: self.export_midpoint_aspects(
                treeview=treeview,
                heading=heading("natal", "transit")
            )
        )
        midpoint_aspects.add_command(
            label=heading("natal", "progressed"),
            command=lambda: self.export_midpoint_aspects(
                treeview=treeview,
                heading=heading("natal", "progressed")
            )
        )
        midpoint_aspects.add_command(
            label=heading("transit", "natal"),
            command=lambda: self.export_midpoint_aspects(
                treeview=treeview,
                heading=heading("transit", "natal")
            )
        )
        midpoint_aspects.add_command(
            label=heading("progressed", "natal"),
            command=lambda: self.export_midpoint_aspects(
                treeview=treeview,
                heading=heading("progressed", "natal")
            )
        )

    def button_3_on_entry(self):
        self.right_click.add_command(
            label="Copy",
            command=lambda: self.master.focus_get(
            ).event_generate('<<Copy>>')
        )
        self.right_click.add_command(
            label="Cut",
            command=lambda: self.master.focus_get(
            ).event_generate('<<Cut>>')
        )
        self.right_click.add_command(
            label="Delete",
            command=lambda: self.master.focus_get(
            ).event_generate('<<Clear>>')
        )
        self.right_click.add_command(
            label="Paste",
            command=lambda: self.master.focus_get(
            ).event_generate('<<Paste>>')
        )

    @staticmethod
    def get_values(values: list = []):
        values = [
            values[1],
            values[2],
            *values[8:],
            ".".join([str(i).zfill(2) for i in values[3:6]])
            + " "
            + ":".join([str(i).zfill(2) for i in values[6:8]])
        ]
        titles = [
            "Name", "Gender", "Latitude", "Longitude", "Natal Date"
        ]
        return values, titles

    def write_title(
            self,
            wb: workbook.Workbook = None,
            ws: worksheet.Worksheet = None,
            titles: list = [],
            values: list = [],
            synastry: bool = False
    ):
        if synastry:
            cols = ["E", "F", "G", "H", 4, 5]
        else:
            cols = ["A", "B", "C", "D", 0, 1]
        for i, j in enumerate(titles):
            if j in ["Natal Date", "Transit Date", "Progressed Date"]:
                ws.merge_range(
                    f"{cols[0]}{i + 1}:{cols[1]}{i + 1}",
                    j,
                    self.format(
                        align="left",
                        wb=wb,
                        bold=True
                    )
                )
                ws.merge_range(
                    f"{cols[2]}{i + 1}:{cols[3]}{i + 1}",
                    values[i],
                    self.format(
                        align="left",
                        wb=wb
                    )
                )
            else:
                ws.merge_range(
                    f"{cols[0]}{i + 1}:{cols[1]}{i + 1}",
                    j,
                    self.format(
                        bold=True,
                        align="left",
                        wb=wb
                    )
                )
                ws.merge_range(
                    f"{cols[2]}{i + 1}:{cols[3]}{i + 1}",
                    values[i],
                    self.format(
                        align="left",
                        wb=wb
                    )
                )

    def _write(
            self,
            wb: workbook.Workbook = None,
            ws: worksheet.Worksheet = None,
            titles: list = [],
            values: list = [],
            synastry: list = [],
            row: int = 0,
            heading: str = "",
    ):
        self.write_title(
            wb=wb, ws=ws, titles=titles, values=values
        )
        if synastry:
            self.write_title(
                wb=wb,
                ws=ws,
                titles=titles,
                values=synastry,
                synastry=True
            )
        if heading in ["Natal Midpoints", "Progressed Midpoints"]:
            letter = "D"
            _heading = heading.split(" ")[0]
        else:
            letter = "L"
            _heading = heading
        ws.merge_range(
            f"A{row}:{letter}{row}",
            _heading,
            self.format(
                wb=wb,
                align="center",
                bold=True
            )
        )
        if heading in ["Natal Midpoints", "Progressed Midpoints"]:
            pass
        else:
            ws.merge_range(
                f"A{row - 13}:D{row - 13}",
                "Orb Factors",
                self.format(
                    wb=wb,
                    align="center",
                    bold=True
                )
            )
            for i, j in enumerate(ASPECTS):
                ws.merge_range(
                    f"A{i + row - 12}:B{i + row - 12}",
                    f"{j}",
                    self.format(
                        wb=wb,
                        align="left",
                        bold=True
                    )
                )
                ws.merge_range(
                    f"C{i + row - 12}:D{i + row - 12}",
                    f"\u00b1 {ASPECTS[j]['orb']}",
                    self.format(
                        align="center",
                        wb=wb,
                    )
                )

    def write_midpoints(
            self,
            row: int = 0,
            values: list = [],
            selection: list = [],
            heading: str = "",
            date: str = "",
            title: str = ""
    ):
        filename = f"{values[1]}_{heading}"
        midpoint = self.find_midpoints(midpoint=selection)
        values, titles = self.get_values(values=values)
        if date and title:
            values.append(date)
            titles.append(title)
        with workbook.Workbook(f"{filename}.xlsx") as wb:
            ws = wb.add_worksheet()
            self._write(
                wb=wb,
                ws=ws,
                titles=titles,
                values=values,
                heading=heading,
                row=row
            )
            c = 0
            for i in ["Midpoint", "Position"]:
                ws.merge_range(
                    f"{ALPHABETA[c]}{row + 1}:{ALPHABETA[c + 1]}{row + 1}",
                    i,
                    self.format(
                        bold=True,
                        align="center",
                        wb=wb
                    )
                )
                c += 2
            row += 2
            for i, mid in enumerate(midpoint):
                deg_sign = convert_degree(mid[-1])
                deg = dd_to_dms(deg_sign[0]).split()
                sign = deg_sign[1]
                deg = deg[0] + sign[:3] + deg[1]
                reformat = [
                    f"{mid[0]}  |  "
                    f"{mid[1]}", deg
                ]
                ws.merge_range(
                    f"{ALPHABETA[0]}{row + i}:{ALPHABETA[1]}{row + i}",
                    reformat[0],
                    self.format(
                        align="center",
                        wb=wb
                    )
                )
                ws.merge_range(
                    f"{ALPHABETA[2]}{row + i}:{ALPHABETA[3]}{row + i}",
                    reformat[1],
                    self.format(
                        align="center",
                        wb=wb
                    )
                )
        showinfo(title="Info", message="Completed.")

    def export_midpoints(
            self,
            treeview: Treeview = None,
            heading: str = "",
            row: int = 0
    ):
        selected = treeview.selection()
        if not selected:
            pass
        else:
            values = treeview.item(selected)["values"]
            record = Record(
                year=values[5],
                month=values[4],
                day=values[3],
                hour=values[6],
                minute=values[7],
                lat=float(values[8]),
                lon=float(values[9])
            )
            if heading == "Natal Midpoints":
                self.write_midpoints(
                    row=row,
                    selection=record.natal(),
                    heading=heading,
                    values=values,
                )
            elif heading == "Progressed Midpoints":
                self.request_date(
                    title="progressed",
                    heading=heading,
                    values=values,
                    row=row,
                    record=record
                )

    def export_midpoint_aspects(
            self,
            treeview: Treeview = None,
            heading: str = ""
    ):
        selected = treeview.selection()
        if not selected:
            pass
        else:
            values = treeview.item(selected)["values"]
            record = Record(
                year=values[5],
                month=values[4],
                day=values[3],
                hour=values[6],
                minute=values[7],
                lat=float(values[8]),
                lon=float(values[9])
            )
            selected_midpoint = heading.split()[4]
            selected_planet = heading.split()[7]
            filename = f"{values[1]}_{selected_midpoint}_{selected_planet}"
            if selected_midpoint == "natal":
                if selected_planet == "natal":
                    self.write(
                        filename=filename,
                        heading=heading,
                        values=values,
                        midpoint=record.natal(),
                        to_planet=record.natal(),
                        row=20
                    )
                elif selected_planet == "transit":
                    self.request_date(
                        title="transit",
                        record=record,
                        filename=filename,
                        heading=heading,
                        values=values
                    )
                else:
                    self.request_date(
                        title="progressed",
                        record=record,
                        filename=filename,
                        heading=heading,
                        values=values
                    )
            elif selected_midpoint == "transit":
                self.request_date(
                    title="transit",
                    record=record,
                    filename=filename,
                    heading=heading,
                    values=values
                )
            else:
                self.request_date(
                    title="progressed",
                    record=record,
                    filename=filename,
                    heading=heading,
                    values=values
                )

    def synastry_button(
            self,
            toplevel: tk.Toplevel = None,
            treeview_1: Treeview = None,
            treeview_2: Treeview = None,
            values1: list = []
    ):
        selected = treeview_2.selection()
        if not selected:
            pass
        else:
            num = treeview_2.item(selected)["values"][0]
            for i in treeview_1.get_children():
                values2 = treeview_1.item(i)["values"]
                if values2[0] == num:
                    record_1 = Record(
                        year=values1[5],
                        month=values1[4],
                        day=values1[3],
                        hour=values1[6],
                        minute=values1[7],
                        lat=float(values1[8]),
                        lon=float(values1[9])
                    ).natal()
                    record_2 = Record(
                        year=values2[5],
                        month=values2[4],
                        day=values2[3],
                        hour=values2[6],
                        minute=values2[7],
                        lat=float(values2[8]),
                        lon=float(values2[9])
                    ).natal()
                    self.write(
                        filename=f"{values1[1]}_{values2[1]}"
                                 f"_Midpoint_Synastry",
                        heading=f"Aspects between midpoints of"
                                f" {values1[1]}'s natal planets"
                                f" and {values2[1]}'s natal planets",
                        midpoint=record_1,
                        to_planet=record_2,
                        values=values1,
                        row=20,
                        synastry=values2
                    )
                    toplevel.destroy()

    def midpoint_synastry(self, treeview: Treeview = None):
        selected = treeview.selection()
        if not selected:
            pass
        else:
            values = treeview.item(selected)["values"]
            toplevel = tk.Toplevel()
            toplevel.title("Select a record")
            toplevel.geometry("200x250")
            toplevel.resizable(width=False, height=False)
            frame = tk.Frame(master=toplevel)
            frame.pack()
            y_scrollbar = tk.Scrollbar(master=frame, orient="vertical")
            y_scrollbar.pack(side="right", fill="y")
            columns = "No", "Name"
            _treeview = Treeview(
                master=frame,
                show="headings",
                columns=[f"#{i + 1}" for i in range(len(columns))],
                height=10,
                selectmode="extended",

            )
            _treeview.pack()
            y_scrollbar.configure(command=_treeview.yview)
            for i, j in enumerate(columns):
                _treeview.column(
                    column=f"#{i + 1}",
                    minwidth=100,
                    width=100,
                    anchor=tk.CENTER
                )
                self.heading(
                    treeview=_treeview,
                    col=i,
                    text=j
                )
            for i, j in enumerate(treeview.get_children()):
                item = treeview.item(j)["values"][:2]
                if item[0] != values[0]:
                    _treeview.insert(
                        parent="",
                        index=i,
                        values=[
                            col for col in item
                        ]
                    )
            button = tk.Button(
                master=toplevel,
                text="Apply",
                command=lambda: self.synastry_button(
                    treeview_1=treeview,
                    treeview_2=_treeview,
                    values1=values,
                    toplevel=toplevel
                )
            )
            button.pack()

    def write(
            self,
            filename: str = "",
            heading: str = "",
            values: list = [],
            midpoint: list = [],
            to_planet: list = [],
            row: int = 0,
            date: str = "",
            synastry: list = []
    ):
        with workbook.Workbook(f"{filename}.xlsx") as wb:
            ws = wb.add_worksheet()
            m = heading.split()[4]
            t = heading.split()[7]
            values, titles = self.get_values(values=values)
            if m == "natal" and t == "transit":
                values.append(date)
                titles.append("Transit Date")
                row += 1
            elif m == "natal" and t == "progressed":
                values.append(date)
                titles.append("Progressed Date")
                row += 1
            elif m == "transit" and t == "natal":
                values.append(date)
                titles.append("Transit Date")
                row += 1
            elif m == "progressed" and t == "natal":
                values.append(date)
                titles.append("Progressed Date")
                row += 1
            if synastry:
                values2, titles2 = self.get_values(synastry)
            else:
                values2 = []
            self._write(
                wb=wb,
                ws=ws,
                titles=titles,
                values=values,
                heading=heading,
                row=row,
                synastry=values2
            )
            self.insert_values(
                n=0, row=row + 1, wb=wb, ws=ws, bold=True
            )
            self.find_aspects(
                midpoint=midpoint,
                to_planet=to_planet,
                wb=wb,
                ws=ws,
                row=row + 2,
                filename=filename,
                values=values,
                titles=titles,
                synastry=synastry,
            )
        showinfo(title="Info", message="Completed.")

    def insert_values(
            self,
            wb: workbook.Workbook = None,
            ws: worksheet.Worksheet = None,
            values: list = [],
            n: int = 0,
            row: int = 0,
            bold: bool = False
    ):
        if not values:
            row_values = ["Midpoint", "Aspect", "Planet", "Orb"]
        else:
            row_values = values
        for i in row_values:
            ws.merge_range(
                f"{ALPHABETA[n]}{row}:{ALPHABETA[n + 2]}{row}",
                i,
                self.format(
                    wb=wb, align="center", bold=bold
                )
            )
            n += 3

    @staticmethod
    def format(
            wb: xlsxwriter.Workbook = None,
            bold: bool = False,
            align: str = "",
            font_name: str = "Arial",
            font_size: int = 11
    ):
        return wb.add_format(
            {
                "bold": bold,
                "align": align,
                "valign": "vcenter",
                "font_name": font_name,
                "font_size": font_size
            }
        )

    @staticmethod
    def find_midpoints(midpoint: list = []):
        for i in midpoint:
            for j in midpoint[midpoint.index(i) + 1:]:
                mid = 0
                if abs(i[2] - j[2]) > 180:
                    mid += (i[2] + 360 + j[2]) / 2
                elif abs(i[2] - j[2]) < 180:
                    mid += (i[2] + j[2]) / 2
                if mid > 360:
                    mid -= 360
                yield [i[0], j[0], mid]

    @staticmethod
    def reformat_aspect(arr: list = []):
        converted = convert_degree(arr[2])
        dms = dd_to_dms(converted[0]).split()
        sign = converted[1][:3]
        return dms[0] + sign + dms[1]

    def find_difference(
            self,
            midpoint: list = [],
            to_planet: list = [],
    ):
        midpoints = self.find_midpoints(midpoint=midpoint)
        result = {}
        for i in midpoints:
            frmt1 = self.reformat_aspect(arr=i)
            for j in to_planet:
                frmt2 = self.reformat_aspect(arr=j)
                degree_diff = 0
                if abs(i[-1] - j[2]) > 180:
                    degree_diff += 360 - abs(i[-1] - j[2])
                elif abs(i[-1] - j[2]) < 180:
                    degree_diff += abs(i[-1] - j[2])
                result[
                    f"{i[0]}/{i[1]} {frmt1}"
                    f" -> {j[0]} {frmt2}"
                ] = degree_diff                
        return result

    @staticmethod
    def write_interpretation(
            root: ET.Element = None,
            values: list = [],
            output: open = None
    ):
        for i, j in enumerate(root):
            if root[i].get("midpoint") == values[0] \
                    and root[i].get("planet") == values[3]:
                output.write(
                    f"{'-' * 79}\n"
                    f"{values[0]} = {values[3]}\n"
                    f"Aspect: {values[2]}\n"
                    f"Author: {root[i][0].get('author')}\n"
                    f"Text: {root[i][0].text}\n"
                )
                output.flush()

    def find_aspects(
            self,
            midpoint: list = [],
            to_planet: list = [],
            values: list = [],
            titles: list = [],
            synastry: list = [],
            wb: workbook.Workbook = None,
            ws: worksheet.Worksheet = None,
            row: int = 0,
            filename: str = ""
    ):
        diff = self.find_difference(
            midpoint=midpoint,
            to_planet=to_planet,
        )
        tree = ET.parse("interpretations.xml")
        root = tree.getroot()
        row = row
        if not synastry:
            output = open(f"{filename}.txt", "w", encoding="utf-8")
            for i, j in zip(values, titles):
                output.write(
                    f"{j}: {i}\n"
                )
                output.flush()
            output.write("\n")
        else:
            output = None
        for i in diff:
            for j in ASPECTS:
                if ASPECTS[j]["degree"] \
                        - dms_to_dd(ASPECTS[j]["orb"]) < diff[i] \
                        < ASPECTS[j]["degree"] \
                        + dms_to_dd(ASPECTS[j]["orb"]):
                    values = i.replace(
                        '->', j
                    ).split() \
                        + [str(ASPECTS[j]['degree'] - diff[i])]
                    values[0] = values[0].replace("/", " | ")
                    dd = float(values[-1])
                    if not synastry:
                        self.write_interpretation(
                            root=root,
                            values=values,
                            output=output
                        )
                    values[-1] = dd_to_dms(abs(dd))
                    values[0] = f"{values[0]}   {values[1]}"
                    values[3] = f"{values[3]}   {values[4]}"
                    values = [
                        values[0], values[2], values[3], values[5]
                    ]
                    if dd < 0:
                        values[-1] = "-" + values[-1]
                    self.insert_values(
                        wb=wb,
                        ws=ws,
                        row=row,
                        values=values
                    )
                    row += 1
        if output:
            output.close()

    def request_date(
            self,
            title: str = "",
            filename: str = "",
            heading: str = "",
            record: Record = None,
            values: list = [],
            row: int = 0
    ):
        toplevel = tk.Toplevel(master=None)
        toplevel.title(title.capitalize())
        toplevel.geometry("200x100")
        toplevel.resizable(width=False, height=False)
        topframe = tk.Frame(master=toplevel, relief="sunken", bd=1)
        topframe.pack(expand=True, fill=tk.Y)
        entries = []
        for i in ["Day", "Month", "Year"]:
            frame = tk.Frame(topframe)
            frame.pack(side="left")
            label = tk.Label(master=frame, text=i, font="Default 11 bold")
            label.pack()
            if i != "Year":
                entry = tk.Entry(master=frame, width=2)
                entry.bind(
                    sequence="<KeyRelease>",
                    func=lambda event: self.max_char(
                        event=event,
                        limit=2
                    )
                )
            else:
                entry = tk.Entry(master=frame, width=4)
                entry.bind(
                    sequence="<KeyRelease>",
                    func=lambda event: self.max_char(
                        event=event,
                        limit=4
                    )
                )
            entry.pack()
            entries.append(entry)
        button = tk.Button(
            master=toplevel,
            text="Apply",
            command=lambda: self.apply_date(
                toplevel=toplevel,
                record=record,
                entries=entries,
                filename=filename,
                values=values,
                heading=heading,
                row=row
            )
        )
        button.pack()

    def apply_date(
            self,
            toplevel: tk.Toplevel = None,
            record: Record = None,
            entries: list = [],
            values: list = [],
            filename: str = "",
            heading: str = "",
            row: int = 0
    ):
        try:
            dt.strptime(
                ".".join([i.get() for i in entries]), "%d.%m.%Y"
            )
            date = [int(i.get()) for i in entries[::-1]]
            _date = ".".join([str(i).zfill(2) for i in date[::-1]]) \
                + " 00:00"
            if heading == "Progressed Midpoints":
                self.write_midpoints(
                    row=row,
                    selection=record.progressed(*date),
                    heading=heading,
                    values=values,
                    title="Progressed Date",
                    date=_date
                )
                toplevel.destroy()
            else:
                midpoint = heading.split()[4]
                to_planet = heading.split()[7]
                if midpoint == "natal" and to_planet == "transit":
                    m = record.natal()
                    t = record.transit(*date)
                elif midpoint == "natal" and to_planet == "progressed":
                    m = record.natal()
                    t = record.progressed(*date)
                elif midpoint == "transit" and to_planet == "natal":
                    m = record.transit(*date)
                    t = record.natal()
                else:
                    m = record.progressed(*date)
                    t = record.natal()
                self.write(
                    filename=filename,
                    heading=heading,
                    values=values,
                    midpoint=m,
                    to_planet=t,
                    row=20,
                    date=_date
                )
                toplevel.destroy()
        except ValueError:
            showinfo(title="Warning", message="Invalid Date.")

    def choose_orb_factor(self):
        toplevel = tk.Toplevel()
        toplevel.title("Orb Factor")
        toplevel.geometry("200x300")
        toplevel.resizable(width=False, height=False)
        default_orbs = [ASPECTS[i]["orb"] for i in ASPECTS]
        orb_entries = []
        frame = tk.Frame(master=toplevel)
        frame.pack()
        for i, j in enumerate(list(ASPECTS.keys())):
            aspect_label = tk.Label(
                master=frame,
                text=f"{j}"
            )
            aspect_label.grid(row=i, column=0, sticky="w")
            orb_entry = tk.Entry(master=frame, width=7)
            orb_entry.grid(row=i, column=1)
            orb_entry.insert(0, default_orbs[i])
            orb_entries.append(orb_entry)
        apply_button = tk.Button(
            master=frame,
            text="Apply",
            command=lambda: self.change_orb_factors(
                orb_entries=orb_entries,
                toplevel=toplevel
            )
        )
        apply_button.grid(row=11, column=0, columnspan=3)

    def view_interpretations(self):
        toplevel = tk.Toplevel()
        toplevel.title("Interpretations")
        toplevel.geometry("640x480")
        toplevel.resizable(width=False, height=False)
        columns = ["Number", "Midpoint", "Planet", "Availability"]
        y_scrollbar = tk.Scrollbar(master=toplevel, orient="vertical")
        y_scrollbar.pack(side="right", fill="y")
        treeview = Treeview(
            master=toplevel,
            show="headings",
            columns=[f"#{i + 1}" for i in range(len(columns))],
            height=20,
            selectmode="extended",
            yscrollcommand=y_scrollbar.set
        )
        treeview.pack(expand=True, fill="both")
        y_scrollbar.configure(command=treeview.yview)
        for i, j in enumerate(columns):
            treeview.column(
                column=f"#{i + 1}",
                minwidth=100,
                width=100,
                anchor=tk.CENTER
            )
            self.heading(
                treeview=treeview,
                col=i,
                text=j
            )
        treeview.bind(
            sequence="<Double-Button-1>",
            func=lambda event: self.open_interpretation_panel(
                treeview=treeview
            )
        )
        self.read_xml(treeview=treeview)

    def open_interpretation_panel(self, treeview: Treeview = None):
        selected = treeview.selection()
        if selected:
            item = treeview.item(selected)["values"][:-1]
            toplevel = tk.Toplevel()
            toplevel.title(" = ".join(item[1:]))
            toplevel.resizable(width=False, height=False)
            tree = ET.parse("interpretations.xml")
            root = tree.getroot()
            _midpoint = root[item[0] - 1].get("midpoint")
            _planet = root[item[0] - 1].get("planet")
            _author = root[item[0] - 1][0].get("author")
            _text = root[item[0] - 1][0].text
            topframe = tk.Frame(master=toplevel, bd=1, relief="sunken")
            topframe.pack()
            bottomframe = tk.Frame(master=toplevel, bd=1, relief="sunken")
            bottomframe.pack()
            label_author = tk.Label(
                master=topframe,
                text="Author",
                font="Default 11 bold"
            )
            label_author.grid(row=0, column=0)
            entry = tk.Entry(master=topframe, width=40)
            entry.grid(row=1, column=0)
            entry.insert(0, _author)
            label_text = tk.Label(
                master=bottomframe,
                text="Text",
                font="Default 11 bold"
            )
            label_text.pack()
            text = tk.Text(master=bottomframe)
            text.pack()
            if not _text:
                pass
            elif _text == "\n":
                text.delete("1.0", tk.END)
            else:
                text.insert("insert", _text)
            entry.bind(
                sequence="<Button-3>",
                func=lambda event: self.popup(
                    event=event,
                    func=self.button_3_on_entry,
                    treeview=None
                )
            )
            entry.bind(
                sequence="<Button-1>",
                func=lambda event: self.popdown()
            )
            entry.bind(
                sequence="<Control-KeyRelease-a>",
                func=lambda event: event.widget.select_range(0, tk.END)
            )
            text.bind(
                sequence="<Button-3>",
                func=lambda event: self.popup(
                    event=event,
                    func=self.button_3_on_entry,
                    treeview=None
                )
            )
            text.bind(
                sequence="<Button-1>",
                func=lambda event: self.popdown()
            )
            text.bind(
                sequence="<Control-KeyRelease-a>",
                func=lambda event: event.widget.tag_add(
                    tk.SEL, "1.0", tk.END
                )
            )
            button = tk.Button(
                master=toplevel,
                text="Apply",
                command=lambda: self.apply_interpretation(
                    root=root,
                    tree=tree,
                    treeview=treeview,
                    toplevel=toplevel,
                    text=text,
                    entry=entry,
                    selected=selected,
                    item=item
                )
            )
            button.pack()

    @staticmethod
    def apply_interpretation(
            tree: ET = None,
            root: ET.Element = None,
            treeview: Treeview = None,
            toplevel: tk.Toplevel = None,
            text: tk.Text = None,
            entry: tk.Entry = None,
            selected: tuple = (),
            item: list = []
    ):
        if text.get("1.0", tk.END) not in ["", "\n"]:
            state = "Available"
            root[item[0] - 1][0].text = text.get("1.0", tk.END)
        else:
            state = "Not Available"
            root[item[0] - 1][0].text = ""
        root[item[0] - 1][0].set("author", entry.get())
        tree.write("interpretations.xml")
        item = treeview.item(selected)["values"][:-1]
        item.append(state)
        treeview.delete(selected)
        treeview.insert(
            parent="",
            index=item[0] - 1,
            values=[
                col for col in item
            ]
        )
        toplevel.destroy()

    @staticmethod
    def read_xml(
            treeview: Treeview = None
    ):
        tree = ET.parse("interpretations.xml")
        root = tree.getroot()
        planets = {i: j for i, j in PLANETS.items()}
        planets.update({"Asc": None, "MC": None})
        for i in range(len(root)):
            midpoint = root[i].get("midpoint")
            planet = root[i].get("planet")
            text = root[i][0].text
            mid1, mid2 = midpoint.split(" | ")
            midpoint = f"{mid1} | {mid2}"
            if text and text:
                availability = "Available"
            else:
                availability = "Not Available"
            values = [i + 1, midpoint, planet, availability]
            treeview.insert(
                parent="",
                index=i,
                values=[
                    col for col in values
                ]
            )

    @staticmethod
    def change_orb_factors(
            orb_entries: list = [],
            toplevel: tk.Toplevel = None
    ):
        error = False
        for i, j in enumerate(ASPECTS):
            if not re.findall(
                    "[0-9]\\u00b0\\s[0-9]*'\\s[0-9]*\"",
                    orb_entries[i].get()
            ):
                error = True
            else:
                ASPECTS[j]["orb"] = orb_entries[i].get()
        if error:
            showinfo(title="Warning", message="Invalid Orb Factor")
        else:
            toplevel.destroy()

    def about(self):
        toplevel = tk.Toplevel()
        toplevel.title("About TkMidpoint")
        name = "TkMidpoint"
        version, _version = "Version:", __version__
        build_date, _build_date = "Built Date:", "04.02.2020"
        update_date, _update_date = "Update Date:", \
            dt.strftime(
                dt.fromtimestamp(os.stat(sys.argv[0]).st_mtime),
                "%d.%m.%Y"
            )
        developed_by, _developed_by = "Developed By:", \
                                      "Tanberk Celalettin Kutlu"
        contact, _contact = "Contact:", "tckutlu@gmail.com"
        github, _github = "GitHub:", \
                          "https://github.com/dildeolupbiten/TkMidpoint"
        tframe1 = tk.Frame(master=toplevel, bd="2", relief="groove")
        tframe1.pack(fill="both")
        tframe2 = tk.Frame(master=toplevel)
        tframe2.pack(fill="both")
        tlabel_title = tk.Label(master=tframe1, text=name, font="Arial 25")
        tlabel_title.pack()
        for i, j in enumerate(
                (
                    version,
                    build_date,
                    update_date,
                    developed_by,
                    contact,
                    github
                )
        ):
            tlabel_info_1 = tk.Label(
                master=tframe2,
                text=j,
                font="Arial 12",
                fg="red"
            )
            tlabel_info_1.grid(row=i, column=0, sticky="w")
        for i, j in enumerate(
                (
                    _version,
                    _build_date,
                    _update_date,
                    _developed_by,
                    _contact,
                    _github
                )
        ):
            if j == _github:
                tlabel_info_2 = tk.Label(
                    master=tframe2,
                    text=j,
                    font="Arial 12",
                    fg="blue",
                    cursor="hand2"
                )
                url1 = "https://github.com/dildeolupbiten/TkMidpoint"
                tlabel_info_2.bind(
                    "<Button-1>",
                    lambda event: self.callback(url1))
            elif j == _contact:
                tlabel_info_2 = tk.Label(
                    master=tframe2,
                    text=j,
                    font="Arial 12",
                    fg="blue",
                    cursor="hand2"
                )
                url2 = "mailto://tckutlu@gmail.com"
                tlabel_info_2.bind(
                    "<Button-1>",
                    lambda event: self.callback(url2))
            else:
                tlabel_info_2 = tk.Label(
                    master=tframe2,
                    text=j,
                    font="Arial 12"
                )
            tlabel_info_2.grid(row=i, column=1, sticky="w")

    @staticmethod
    def update_script():
        url_1 = "https://raw.githubusercontent.com/dildeolupbiten/" \
                "TkMidpoint/master/TkMidpoint.py"
        url_2 = "https://raw.githubusercontent.com/dildeolupbiten/" \
                "TkMidpoint/master/README.md"
        data_1 = urllib.request.urlopen(
            url=url_1,
            context=ssl.SSLContext(ssl.PROTOCOL_SSLv23)
        )
        data_2 = urllib.request.urlopen(
            url=url_2,
            context=ssl.SSLContext(ssl.PROTOCOL_SSLv23)
        )
        with open(
                file="TkMidpoint.py",
                mode="r",
                encoding="utf-8"
        ) as f:
            var_1 = [i.decode("utf-8") for i in data_1]
            var_2 = [i.decode("utf-8") for i in data_2]
            var_3 = [i for i in f]
            if var_1 == var_3:
                showinfo(
                    title="Update",
                    message="Program is up-to-date."
                )
            else:
                with open(
                        file="README.md",
                        mode="w",
                        encoding="utf-8"
                ) as g:
                    for i in var_2:
                        g.write(i)
                        g.flush()
                with open(
                        file="TkMidpoint.py",
                        mode="w",
                        encoding="utf-8"
                ) as h:
                    for i in var_1:
                        h.write(i)
                        h.flush()
                    showinfo(
                        title="Update",
                        message="Program is updated."
                    )
                    if os.name == "posix":
                        subprocess.Popen(
                            ["python3", "TkMidpoint.py"]
                        )
                        import signal
                        os.kill(os.getpid(), signal.SIGKILL)
                    elif os.name == "nt":
                        subprocess.Popen(
                            ["python", "TkMidpoint.py"]
                        )
                        os.system(f"TASKKILL /F /PID {os.getpid()}")

    @staticmethod
    def callback(url):
        webbrowser.open_new(url)


def main():
    root = tk.Tk()
    root.title("TkMidpoint")
    root.geometry(
        f"{root.winfo_screenwidth()}x{root.winfo_screenheight()}"
    )
    App(master=root).mainloop()


if __name__ == "__main__":
    main()