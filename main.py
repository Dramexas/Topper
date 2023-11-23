import tkinter as tk
import win32gui
import win32con
import win32process
import psutil

def list_visible_windows():
    def window_enum_handler(hwnd, resultList):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) != '':
            threadid, pid = win32process.GetWindowThreadProcessId(hwnd)
            process = psutil.Process(pid)
            resultList.append((hwnd, win32gui.GetWindowText(hwnd), process.name()))
    windows = []
    win32gui.EnumWindows(window_enum_handler, windows)
    return [window for window in windows if window[1]]

def create_window():
    root = tk.Tk()
    root.title("Topper")  # Set the window title

    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    window_width = screen_width // 3  # 1/3 of the screen width
    window_height = screen_height // 10  # 1/10 of the screen height

    frame = tk.Frame(root)
    frame.pack()

    root.geometry(f"{window_width}x{window_height}")  # Set the window size

    root.configure(bg='white')  # Set the background color to white

    listbox = tk.Listbox(root, bg='white', fg='black', bd=0, highlightthickness=0, selectbackground='darkgray')  # Set the listbox style 
    listbox.pack(fill=tk.BOTH, expand=1)

    # Create a dictionary to store the state of the option for each window
    option_states = {}

    context_menu = tk.Menu(root, tearoff=0)

    def toggle_option():
        # Get the selected item
        selected = listbox.get(listbox.curselection())

        # Toggle the state of the option for this item
        if selected in option_states and option_states[selected]:
            print(f"Option 1 deactivated for window {selected}")
            option_states[selected] = False
        else:
            print(f"Option 1 activated for window {selected}")
            option_states[selected] = True

            # Bring the window to the foreground
            hwnd = next((window[0] for window in windows if window[1] == selected), None)
            if hwnd is not None:
                parent_hwnd = win32gui.GetParent(hwnd)
                if parent_hwnd != 0:
                    win32gui.ShowWindow(parent_hwnd, win32con.SW_SHOWMINIMIZED)
                    win32gui.ShowWindow(parent_hwnd, win32con.SW_RESTORE)
                win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)

        # Update the context menu
        context_menu.entryconfig(0, label=f"Option 1 {'(activated)' if option_states[selected] else ''}")

    context_menu.add_command(label="Option 1", command=toggle_option)

    def show_context_menu(event):
        # Select the item under the cursor
        listbox.selection_clear(0, tk.END)
        listbox.selection_set(listbox.nearest(event.y))

        # Update the context menu
        selected = listbox.get(listbox.curselection())
        context_menu.entryconfig(0, label=f"Option 1 {'(activated)' if selected in option_states and option_states[selected] else ''}")

        # Show the context menu
        context_menu.post(event.x_root, event.y_root)

    listbox.bind("<Button-3>", show_context_menu)

    # Add windows to the listbox
    windows = list_visible_windows()
    for hwnd, window_title, process_name in windows:
        listbox.insert(tk.END, window_title)
    
    def update_windows():
        # Clear the listbox
        listbox.delete(0, tk.END)

        # Add windows to the listbox
        windows = list_visible_windows()
        windows.sort(key=lambda window: window[1])  # Sort the windows by title

        for hwnd, window_title, process_name in windows:
            # Check if the window is still open
            if win32gui.IsWindow(hwnd):
                listbox.insert(tk.END, window_title)
                if window_title not in option_states:
                    option_states[window_title] = False

                # Check if the window is minimized
                if win32gui.IsIconic(hwnd):
                    # Reset the option for this window
                    option_states[window_title] = False
        root.after(5000, update_windows)

    root.after(5000, update_windows)


    root.mainloop()

create_window()