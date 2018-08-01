import random
from tkinter import *
import glob
import serial
import queue
import threading
import time


class GuiPart:

    label = 'Undefined'
    ser = serial.Serial()
    ser.baudrate = 9600

    def __init__(self, master):
        global label
        # Set up the GUI
        master.resizable(False, False)
        master.title("Roll-Off Roof")
        mf = Frame(master, height=245, width=250)
        mf.pack_propagate(0)  # don't shrink
        mf.pack()

        menubar = Menu(master)
        connectmenu = Menu(menubar, tearoff=0)

        portlist = Menu(connectmenu, tearoff=0)
        ports = self.get_serial_ports_list()
        if not ports:
            portlist.add_command(label='null')
        else:
            for port in ports:
                portlist.add_command(label=port, command=lambda: self.select_port(port))

        connectmenu.add_cascade(label='Select port', menu=portlist, underline=0)
        menubar.add_cascade(label="Connect", menu=connectmenu)
        connectmenu.add_command(label='Connect', command=lambda: self.connect_port(port_selected))
        connectmenu.add_command(label='Disconnect', command=lambda: self.disconnect_port())
        master.config(menu=menubar, bd=5)

        label_frame = Frame(master, height=30, bg='red')
        connection_label = Label(label_frame, text="Port connection state:", font=("Helvetica", 11))
        connection_label.pack(side=LEFT)
        self.port_state_label = Label(label_frame, text=self.label, font=("Helvetica", 11))
        self.port_state_label.pack(side=LEFT)
        label_frame.place(x=0, y=0)

        # open button
        open_btn_frame = Frame(master, height=40, width=80)
        open_btn_frame.pack_propagate(0)  # don't shrink
        open_btn_frame.pack()
        open_btn_frame.place(x=10, y=30)
        open_button = Button(open_btn_frame, text='Open', command=lambda: self.open_roof())
        open_button.pack(fill=BOTH, expand=1)

        # Close button
        close_btn_frame = Frame(master, height=40, width=80)
        close_btn_frame.pack_propagate(0)  # don't shrink
        close_btn_frame.pack()
        close_btn_frame.place(x=10, y=80)
        close_button = Button(close_btn_frame, text='Close', command=lambda: self.close_roof())
        close_button.pack(fill=BOTH, expand=1)

        # Stop button
        stop_btn_frame = Frame(master, height=40, width=80)
        stop_btn_frame.pack_propagate(0)  # don't shrink
        stop_btn_frame.pack()
        stop_btn_frame.place(x=10, y=130)
        stop_button = Button(stop_btn_frame, text='Stop', command=lambda: self.stop_roof())
        stop_button.pack(fill=BOTH, expand=1)

        roof_state_frame = Frame(master, height=80, width=120)
        roof_state_l1 = Label(roof_state_frame, text='Roof state:', font=("Helvetica", 12))
        roof_state_l1.pack(fill=X)
        self.roof_state_l2 = Label(roof_state_frame, text='Undefined', font=("Helvetica", 12))
        self.roof_state_l2.pack(fill=X)
        roof_state_frame.pack_propagate(0)
        roof_state_frame.place(x=112, y=60)

        # Separator
        separator = Frame(height=3, width=250, bd=1, relief=SUNKEN)
        separator.place(x=0, y=180)

        # Rails heating status
        rails_heat_frame = Frame(master, height=30, width=250)
        rails_heat_label = Label(rails_heat_frame, text='Rails heating', font=("Helvetica", 12))
        rails_heat_label.pack()
        rails_heat_frame.pack_propagate(0)
        rails_heat_frame.place(x=0, y=182)

        # On/Off buttons
        heat_on_btn_frame = Frame(master, height=30, width=50)
        heat_on_btn_frame.pack_propagate(0)  # don't shrink
        heat_on_btn_frame.pack()
        heat_on_btn_frame.place(x=10, y=210)
        heat_on_btn = Button(heat_on_btn_frame, text='On', command=lambda: self.heat_on())
        heat_on_btn.pack(fill=BOTH, expand=1)

        heat_off_btn_frame = Frame(master, height=30, width=50)
        heat_off_btn_frame.pack_propagate(0)  # don't shrink
        heat_off_btn_frame.pack()
        heat_off_btn_frame.place(x=70, y=210)
        heat_off_btn = Button(heat_off_btn_frame, text='Off', command=lambda: self.heat_off())
        heat_off_btn.pack(fill=BOTH, expand=1)

        heat_state_frame = Frame(master, height=40, width=100)
        heat_state_l1 = Label(heat_state_frame, text='Heat state:', font=("Helvetica", 10))
        heat_state_l1.pack(side=LEFT)
        self.heat_state_l2 = Label(heat_state_frame, text='N/A', font=("Helvetica", 10))
        self.heat_state_l2.pack(side=LEFT)
        heat_state_frame.pack_propagate(0)
        heat_state_frame.place(x=140, y=205)

    def get_serial_ports_list(self):
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

    port_selected = ''

    def select_port(self, port):
        global port_selected
        print('Port: %s' % port)
        port_selected = port
        print(port_selected)
        return port_selected

    def connect_port(self, p):
        self.ser.port = p
        self.ser.open()
        if self.ser.is_open:
            self.show_connection_state('Connected')
            print('port is connected')

            # ser.write('HANDSHAKE#'.encode(encoding='UTF-8'))
            self.request_roof_state()
            #time.sleep(1)
            self.request_heating_state()
        else:
            self.show_connection_state('Disconnected')
        return self.ser


    def disconnect_port(self):
        #self.ser.close()
        if self.ser.is_open:
            self.show_connection_state('Connected')
        else:
            self.show_connection_state('Disconnected')
            self.upd_labels(self.roof_state_l2, 'Undefined', 'black')
            self.upd_labels(self.heat_state_l2, 'N/A', 'black')

    def show_connection_state(self, l):
        global label
        if l == 'Connected':
            label = 'Connected'
            self.port_state_label.configure(text=label, fg="green")
        elif l == 'Disconnected':
            label = 'Disconnected'
            self.port_state_label.configure(text=label, fg="red")
        else:
            label = 'Undefined'

    def request_roof_state(self):
        request = 'GETSTATE#'
        self.ser.write(request.encode(encoding='UTF-8'))
        print(request.encode(encoding='UTF-8'))

    def request_heating_state(self):
        request = 'HEATSTATE#'
        self.ser.write(request.encode(encoding='UTF-8'))

    def open_roof(self):
        self.ser.write('OPEN#'.encode(encoding='UTF-8'))
        print('open')
        # self.request_roof_state()

    def close_roof(self):
        self.ser.write('CLOSE#'.encode(encoding='UTF-8'))
        print('close')
        # self.request_roof_state()

    def stop_roof(self):
        self.ser.write('STOP#'.encode(encoding='UTF-8'))
        print('stop')
        # request_roof_state()

    def heat_on(self):
        self.ser.write('HEATON#'.encode(encoding='UTF-8'))
        print('heat on')
        # request_heating_state()

    def heat_off(self):
        self.ser.write('HEATOFF#'.encode(encoding='UTF-8'))
        print('heat off')
        # request_heating_state()

    def upd_labels(self, label, state, color):
        label.configure(text=state, fg=color)



class ThreadedClient:
    """
    Launch the main part of the GUI and the worker thread. periodicCall and
    endApplication could reside in the GUI part, but putting them here
    means that you have all the thread controls in a single place.
    """
    def __init__(self, master):
        """
        Start the GUI and the asynchronous threads. We are in the main
        (original) thread of the application, which will later be used by
        the GUI as well. We spawn a new thread for the worker (I/O).
        """
        self.master = master

        # Create the queue
        self.q = queue.Queue()

        # Set up the GUI part
        self.gui = GuiPart(master)

        # Set up the thread to do asynchronous I/O
        # More threads can also be created and used, if necessary
        self.running = 1

        self.thread1 = threading.Thread(target=self.workerThread1)
        self.thread1.daemon = True
        self.thread1.start()

        # Start the periodic call in the GUI to check if the queue contains
        # anything
        self.periodicCall()

    def process_queue(self):
        while self.q.qsize():
            try:
                msg = self.q.get()
                print('from queue msg - %s' % msg)
                if msg == '0':
                    state = 'Open'
                    color = 'green'
                    self.gui.upd_labels(self.gui.roof_state_l2, state, color)
                elif msg == '1':
                    state = 'Closed'
                    color = 'black'
                    self.gui.upd_labels(self.gui.roof_state_l2, state, color)
                elif msg == '2':
                    state = 'Opening'
                    color = 'black'
                    self.gui.upd_labels(self.gui.roof_state_l2, state, color)
                elif msg == '3':
                    state = 'Closing'
                    color = 'black'
                    self.gui.upd_labels(self.gui.roof_state_l2, state, color)
                elif msg == '4':
                    state = 'Stop'
                    color = 'black'
                    self.gui.upd_labels(self.gui.roof_state_l2, state, color)
                elif msg == '5':
                    state = 'On'
                    color = 'green'
                    self.gui.upd_labels(self.gui.heat_state_l2, state, color)
                elif msg == '6':
                    state = 'Off'
                    color = 'black'
                    self.gui.upd_labels(self.gui.heat_state_l2, state, color)
                elif msg == '7':
                    state = 'Error'
                    color = 'red'
                    self.gui.upd_labels(self.gui.roof_state_l2, state, color)

            except:
                # just on general principles, although we don't
                # expect this branch to be taken in this case
                pass

    def periodicCall(self):
        """
        Check every 200 ms if there is something new in the queue.
        """
        self.process_queue()

        self.master.after(200, self.periodicCall)

    def workerThread1(self):
        """
        This is where we handle the asynchronous I/O. For example, it may be
        a 'select(  )'. One important thing to remember is that the thread has
        to yield control pretty regularly, by select or otherwise.
        """
        while self.running:
            if self.gui.ser.is_open:

                if self.gui.ser.inWaiting():
                    text = self.gui.ser.readline(self.gui.ser.inWaiting())
                    text = serial.unicode(text, errors='ignore')
                    print('serial received - %s' % text)
                    self.q.put(text)
                    self.process_queue()

    def endApplication(self):
        self.running = 0

master = Tk()

client = ThreadedClient(master)
master.mainloop()