import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from scipy.signal import medfilt
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def main():
    # Load new data?
    load_new_data = input("Load new data? (yes/no): ").strip().lower()

    if load_new_data == 'yes':
        # Clear variables
        global data, filename, path
        
        # Select Excel file
        Tk().withdraw()  # Hide the main Tkinter window
        filename = askopenfilename(filetypes=[("Excel files", "*.xlsx")], title="Select xlsx file")
        
        if filename:
            data = pd.read_excel(filename)
        else:
            print("No file selected.")
            return
    else:
        # Your code for the "No" case, if needed
        return

    f_max = 1_000_000  # max. frequency
    f_max_index = 0

    # Convert data to numpy array
    data_array = data.to_numpy()
    m, n = data_array.shape
    anzahl_elektroden = n // 3

    frequenz = data_array[:, 0]

    for i in range(len(frequenz)):
        if frequenz[i] > f_max:
            f_max_index = i
            break

    if f_max_index > 0:
        data_array = data_array[f_max_index+1:, :]

    plt.figure(figsize=(12, 7))

    plt.subplot(211)
    plt.title('Amplitude vs Frequency')
    for i in range(anzahl_elektroden):
        frequenz = data_array[:, 3 * i]
        amplitude = data_array[:, 3 * i + 1]
        amplitude_filtered = medfilt(amplitude, kernel_size=7)  # Median Filtered with kernel size 7
        
        plt.plot(frequenz, amplitude_filtered)
        plt.xscale('log')
        plt.yscale('log')
        plt.xlim([10, f_max])
    
    plt.ylabel('Z [Ω]')
    plt.grid(True)

    plt.subplot(212)
    plt.title('Phase vs Frequency')
    for i in range(anzahl_elektroden):
        frequenz = data_array[:, 3 * i]
        phase = data_array[:, 3 * i + 2]
        phase_filtered = medfilt(phase, kernel_size=7)  # Median Filtered with kernel size 7
        
        plt.plot(frequenz, phase_filtered)
        plt.xscale('log')
        plt.xlim([10, f_max])
    
    plt.xlabel('f [Hz]')
    plt.ylabel('φ [°]')
    plt.grid(True)

    plt.tight_layout()
    plt.show()

if __name__ == "__main__":
    main()
