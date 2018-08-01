from tkinter import *
from tkinter import ttk
import glob
import serial
import queue
import threading
import win32com.client
import time
from queue import Empty

window = None

label = 'Undefined'
ser = serial.Serial()
ser.baudrate = 9600


class SerialThread(threading.Thread):
    def __init__(self, queue):
        threading.Thread.__init__(self)
        self.queue = queue
        self._stop_event = threading.Event()

    def run(self):
        global ser
        while not self.stopped():
            if ser.is_open:
                text = ser.read()
                text = serial.unicode(text, errors='ignore')
                print('serial received - %s' % text)
                self.queue.put(text)
                process_queue()
            else:
                self.stop()

    def stop(self):
        self._stop_event.set()

    def stopped(self):
        return self._stop_event.is_set()


def get_serial_ports_list():
    if sys.platform.startswith('win'):
        ports = ['COM%s' % (i + 1) for i in range(256)]
    elif sys.platform.startswith('linux') or sys.platform.startswith('cygwin'):
        ports = glob.glob('/dev/tty[A-Za-z]*')
    elif sys.platform.startswith('darwin'):
        ports = glob.glob('/dev/tty.*')
    else:
        raise EnvironmentError('Unsupported platform')

    result = []
    for port in ports:
        try:
            s = serial.Serial(port)
            s.close()
            result.append(port)
        except (OSError, serial.SerialException):
            pass
    return result

ports = get_serial_ports_list()
if not ports:
    ports = ['null']


port_selected = ''


def select_port(port):
    global port_selected
    print('Port: %s' % port)
    port_selected = port
    print(port_selected)
    return port_selected


def connect_port(p):
    global ser
    ser.port = p
    ser.open()
    if ser.is_open:
        show_connection_state('Connected')
        print('port is connected')
        thread = SerialThread(queue)
        thread.daemon = True
        thread.start()
        #time.sleep(1) #Sleep is required for boards on CH340 chip
        request_roof_state()
        time.sleep(1)
        request_heating_state()
        time.sleep(1)
        request_roof_position()
        time.sleep(1)
    else:
        show_connection_state('Disconnected')
    return ser


def test_connection(p):
    global autoconnect_stop
    ser.port = p
    ser.open()
    if ser.is_open:
        ser.write('HANDSHAKE#'.encode(encoding='UTF-8'))
        master.after(100)
        if ser.inWaiting():
            if ser.readline(ser.inWaiting()).decode("utf-8") == 'HANDSHAKE':
                autoconnect_stop = True
                ser.close()
                connect_port(p)
                pass
            else:
                ser.close()
        else:
            ser.close()


def disconnect_port():
    ser.close()
    if ser.is_open:
        show_connection_state('Connected')
    else:
        show_connection_state('Disconnected')
        upd_labels(roof_state_l2, 'Undefined', 'black')
        upd_labels(heat_state_l2, 'N/A', 'black')


def show_connection_state(l):
    global label
    if l == 'Connected':
        label = 'Connected'
        port_state_label.configure(text=label, fg="green")
    elif l == 'Disconnected':
        label = 'Disconnected'
        port_state_label.configure(text=label, fg="red")
    else:
        label = 'Undefined'


def request_roof_state():
    request = 'GETSTATE#'
    ser.write(request.encode(encoding='UTF-8'))


def request_heating_state():
    request = 'HEATSTATE#'
    ser.write(request.encode(encoding='UTF-8'))


def request_roof_position():
    request = 'ROOFPOS#'
    ser.write(request.encode(encoding='UTF-8'))


def open_roof():
    ser.write('OPEN#'.encode(encoding='UTF-8'))
    print('open')
    # self.request_roof_state()


def close_roof():
    print('force val - %s' % force_var.get())
    if force_var.get() == 0 and (not telescope_get_park_state()):
        telescope_park()
    ser.write('CLOSE#'.encode(encoding='UTF-8'))
    print('close')
    # self.request_roof_state()


def stop_roof():
    ser.write('STOP#'.encode(encoding='UTF-8'))
    print('stop')


def heat_on():
    ser.write('HEATON#'.encode(encoding='UTF-8'))
    print('heat on')


def heat_off():
    ser.write('HEATOFF#'.encode(encoding='UTF-8'))
    print('heat off')


def process_queue():
    global queue
    while queue.qsize():
        try:
            msg = queue.get()
            print('from queue msg - %s' % msg)
            if msg == '0':
                state = 'Open'
                color = 'green'
                upd_labels(roof_state_l2, state, color)
            elif msg == '1':
                state = 'Closed'
                color = 'black'
                upd_labels(roof_state_l2, state, color)
            elif msg == '2':
                state = 'Opening'
                color = 'black'
                upd_labels(roof_state_l2, state, color)
            elif msg == '3':
                state = 'Closing'
                color = 'black'
                upd_labels(roof_state_l2, state, color)
            elif msg == '4':
                state = 'Stop'
                color = 'black'
                upd_labels(roof_state_l2, state, color)
            elif msg == '5':
                state = 'On'
                color = 'green'
                upd_labels(heat_state_l2, state, color)
            elif msg == '6':
                state = 'Off'
                color = 'black'
                upd_labels(heat_state_l2, state, color)
            elif msg == '7':
                state = 'Error'
                color = 'red'
                upd_labels(roof_state_l2, state, color)
            elif msg.startswith("p"):
                roof_progress = (int(msg[1:])*100)/300 # CHANGE 300 TO THE NUMBER OF TEETH ON TOOTH RACK
                color = 'black'
                upd_labels(roof_state_l2, "%i" % roof_progress + "%", color)
                roof_progress_bar["value"] = roof_progress

        except Empty:
            # just on general principles, although we don't
            # expect this branch to be taken in this case
            pass


def upd_labels(label, state, color):
    label.configure(text=state, fg=color)


def connect_telescope():
    global telescope
    telescope = win32com.client.Dispatch('EQMOD.Telescope')
    telescope.Connected = True
    print('Connected telescope')


def disconnect_telescope():
    global telescope
    telescope.Connected = False
    print('Disconnected telescope')


def telescope_get_park_state():
    global telescope
    connect_telescope()
    park_pos = telescope.AtPark
    disconnect_telescope()
    return park_pos


def set_park_position():
    global telescope
    connect_telescope()
    telescope.SetPark()
    disconnect_telescope()


def telescope_park():
    global telescope
    connect_telescope()
    telescope.Park()
    while not telescope.AtPark:
        time.sleep(2)

# def create_window():
#     global window
#
#     if (window is not None) and window.winfo_exists():
#         window.focus_force()
#     else:
#         ports = get_serial_ports_list()
#         if not ports:
#             ports = ['null']
#
#         window = Toplevel(master)
#         window.resizable(False, False)
#         window.title("Properties")
#         prop_frame = Frame(window, height=200, width=200)
#         prop_frame.pack_propagate(0)  # don't shrink
#         prop_frame.pack()
#
#


master = Tk()
force_var = IntVar()
force_var.set(1)
master.resizable(False, False)
master.title("Roll-Off Roof")
mf = Frame(master, height=275, width=250)
mf.pack_propagate(0)  # don't shrink
mf.pack()

n = ttk.Notebook(mf)
f1 = ttk.Frame(n)   # first page, which would get widgets gridded into it
f2 = ttk.Frame(n)   # second page
f3 = ttk.Frame(n)
n.add(f1, text='Control')
n.add(f2, text='Properties')
n.add(f3, text="Sensors")
n.pack(expand=1, fill="both")

menubar = Menu(master)
connectmenu = Menu(menubar, tearoff=0)


# menubar.add_cascade(label="Menu", menu=connectmenu)
# connectmenu.add_command(label='Properties', command=lambda: create_window())
# master.config(menu=menubar, bd=5)

label_frame = Frame(f1, height=30, bg='red')
connection_label = Label(label_frame, text="Port connection state:", font=("Helvetica", 11))
connection_label.pack(side=LEFT)
port_state_label = Label(label_frame, text=label, font=("Helvetica", 11))
port_state_label.pack(side=LEFT)
label_frame.place(x=0, y=0)

# open button
open_btn_frame = Frame(f1, height=40, width=80)
open_btn_frame.pack_propagate(0)  # don't shrink
open_btn_frame.pack()
open_btn_frame.place(x=10, y=30)
open_button = Button(open_btn_frame, text='Open', command=lambda: open_roof())
open_button.pack(fill=BOTH, expand=1)

# Close button
close_btn_frame = Frame(f1, height=40, width=80)
close_btn_frame.pack_propagate(0)  # don't shrink
close_btn_frame.pack()
close_btn_frame.place(x=10, y=80)
close_button = Button(close_btn_frame, text='Close', command=lambda: close_roof())
close_button.pack(fill=BOTH, expand=1)

# Stop button
stop_btn_frame = Frame(f1, height=40, width=80)
stop_btn_frame.pack_propagate(0)  # don't shrink
stop_btn_frame.pack()
stop_btn_frame.place(x=10, y=130)
stop_button = Button(stop_btn_frame, text='Stop', command=lambda: stop_roof())
stop_button.pack(fill=BOTH, expand=1)

roof_state_frame = Frame(f1, height=100, width=120)
roof_state_l1 = Label(roof_state_frame, text='Roof state:', font=("Helvetica", 12))
roof_state_l1.pack(fill=X)
roof_state_l2 = Label(roof_state_frame, text='Undefined', font=("Helvetica", 12))
roof_state_l2.pack(fill=X)
roof_state_l3 = Label(roof_state_frame, text="0%", font=("Helvetica", 12))
roof_state_l3.pack(fill=X)
roof_state_frame.pack_propagate(0)
roof_state_frame.place(x=112, y=50)

#roof_progress_frame = Frame(master, height=22, width=120)
roof_progress_bar = ttk.Progressbar(roof_state_frame, orient="horizontal", length=120, mode="determinate")
roof_progress_bar.pack(fill=X)
roof_progress_bar["value"] = 0
roof_progress_bar["maximum"] = 100
#roof_progress_frame.pack_propagate(0)
#roof_progress_frame.place(x=112, y=135)

# Separator
separator = Frame(f1, height=3, width=250, bd=1, relief=SUNKEN)
separator.place(x=0, y=180)

# Rails heating status
rails_heat_frame = Frame(f1, height=30, width=250)
rails_heat_label = Label(rails_heat_frame, text='Rails heating', font=("Helvetica", 12))
rails_heat_label.pack()
rails_heat_frame.pack_propagate(0)
rails_heat_frame.place(x=0, y=182)

# On/Off buttons
heat_on_btn_frame = Frame(f1, height=30, width=50)
heat_on_btn_frame.pack_propagate(0)  # don't shrink
heat_on_btn_frame.pack()
heat_on_btn_frame.place(x=10, y=210)
heat_on_btn = Button(heat_on_btn_frame, text='On', command=lambda: heat_on())
heat_on_btn.pack(fill=BOTH, expand=1)

heat_off_btn_frame = Frame(f1, height=30, width=50)
heat_off_btn_frame.pack_propagate(0)  # don't shrink
heat_off_btn_frame.pack()
heat_off_btn_frame.place(x=70, y=210)
heat_off_btn = Button(heat_off_btn_frame, text='Off', command=lambda: heat_off())
heat_off_btn.pack(fill=BOTH, expand=1)

heat_state_frame = Frame(f1, height=40, width=100)
heat_state_l1 = Label(heat_state_frame, text='Heat state:', font=("Helvetica", 10))
heat_state_l1.pack(side=LEFT)
heat_state_l2 = Label(heat_state_frame, text='N/A', font=("Helvetica", 10))
heat_state_l2.pack(side=LEFT)
heat_state_frame.pack_propagate(0)
heat_state_frame.place(x=140, y=205)


#PROPERTIES TAB

port_label = Label(f2, text="Select Port:", font=("Helvetica", 12))
port_label.place(x=17, y=15)

var = StringVar(f2)
var.set(ports[0])  # initial value

# Selec port drop down
option_frame = Frame(f2, height=30, width=100)
option_frame.pack_propagate(0)  # don't shrink
option_frame.pack()
option_frame.place(x=10, y=40)
option = OptionMenu(option_frame, var, *ports)
option.pack(fill=BOTH)

# Connect button
conn_btn_frame = Frame(f2, height=30, width=80)
conn_btn_frame.pack_propagate(0)  # don't shrink
conn_btn_frame.pack()
conn_btn_frame.place(x=150, y=10)
connect_btn = Button(conn_btn_frame, text='Connect', command=lambda: connect_port(var.get()))
connect_btn.pack(fill=BOTH, expand=1)

# Disconnect button
disconn_btn_frame = Frame(f2, height=30, width=80)
disconn_btn_frame.pack_propagate(0)  # don't shrink
disconn_btn_frame.pack()
disconn_btn_frame.place(x=150, y=50)
disconnect_btn = Button(disconn_btn_frame, text='Disconnect', command=lambda: disconnect_port())
disconnect_btn.pack(fill=BOTH, expand=1)

# Separator
separator1 = Frame(f2, height=3, width=250, bd=1, relief=SUNKEN)
separator1.place(x=0, y=90)

# Force roof closing

force_checkbox = Checkbutton(f2, text="Force roof closing while telescope is not parked", justify=LEFT, wraplength=230,
                             variable=force_var)
force_checkbox.place(x=5, y=95)

# Separator
separator2 = Frame(f2, height=3, width=250, bd=1, relief=SUNKEN)
separator2.place(x=0, y=140)

set_park_btn_frame = Frame(f2, height=30, width=80)
set_park_btn_frame.pack_propagate(0)  # don't shrink
set_park_btn_frame.pack()
set_park_btn_frame.place(x=10, y=150)
set_park_btn = Button(set_park_btn_frame, text='Set Park', command=lambda: set_park_position())
set_park_btn.pack(fill=BOTH, expand=1)

set_park_label = Label(f2, text="Set telescope park position")
set_park_label.place(x=95, y=155)

queue = queue.Queue()
mainloop()


