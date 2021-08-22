"""
 Code for data collection, derivation of HR/HRV parameters and analysis of Samsung gear band S3 to extract feature set
 Created by CSIR-CEERI
 Dr. Madan Kumar Lakshmanan with fixes by Dr. Bala Pesala
 Version 4
 Date - 11-December-2019

Important: N-back
 - Approach1 Labels generated as follows:
     * Time 5-15 minutes - Control/Normal data
     * Last 10 minutes - Fatigue data

 Steps to capture and analyse PPG data
 -------------------------------------
 - Read Gearband data from INput Notepad file
 - Bin it to 10 seconds
 - Detrending, smoothing and filtering
 - Pulse Quality Index (PQI) measurement - Time & frequency domain metrics
 - Peak detection
 - Derivation of T.D. & F.D. HR/HRV measures
 - Feature set extraction - 3 minutes time duration with 10 seconds interval

"""

import numpy as np
from scipy.signal import savgol_filter,filtfilt,welch,butter
import seaborn as sns
sns.set(font_scale=12)
sns.set_context("paper") #poster, talk,  paper
from scipy import interpolate


# For colour prinitng of warning/error
# Madan, 8-Nov-2018
import sys

try: color = sys.stdout.shell
except AttributeError: raise RuntimeError("Use IDLE")


import matplotlib.pyplot as plt
plt.rcParams.update({'font.size': 20})

import scipy as sc
from numpy import interp
import time
import peakutils
from itertools import islice
import math
import xlwt
import xlrd


#  FILTERS

def butter_lowpass_filter(data, lowcut, fs, order=5):
    nyq = 0.5 * fs
    low = lowcut / nyq
    #high = highcut / nyq
    b, a = butter(order, [low], btype='low')
    y = filtfilt(b, a, data)
    return y


#Data Conditioning
data=[]
smoothdata=[]
filtereddata=[]
data_trendRemoved=[]

def PlotSignals():
    #plt.clear()

    plt.gca().set_title('PPG Pulses')

    Start=10
    End=200

    lenData=len(filtereddata[Start:End])
    ts=1/fs
    xAxis=np.linspace(0,ts*lenData,lenData)
    plt.plot(xAxis,filtereddata[Start:End],'b')
    plt.xlabel('Time (s)')
    plt.ylabel('Intensity (AU)')
    #plt.show()
def split_overlap(array, size, overlap):
    result = []
    while True:
        if len(array) <= size:
            result.append(array)
            return result
        else:
            result.append(array[:size])
            array = array[size - overlap:]
def hampel(x,k, t0=3):
    n = len(x)
    y = x #y is the corrected series
    L = 1.4826
    for i in range((k + 1),(n - k)):
        if np.isnan(x[(i - k):(i + k+1)]).all():
            continue
        x0 = np.nanmedian(x[(i - k):(i + k+1)])
        S0 = L * np.nanmedian(np.abs(x[(i - k):(i + k+1)] - x0))
        if (np.abs(x[i] - x0) > t0 * S0):
            y[i] = x0
    return(y)

# Data Acquisition

index = 0
rr_peaks=[0]
rr_Biglist=[]
rr_interval=[]
rr_interval_15secs=[]
lastpeakfound=0
flat_list=[]
rr_BigDatalist=[]
correctionFactor=0
lastpeakfound=0
bi=1

CANDIDATE_NAME = "ACHU"
durationWindow = 10 # 10 seconds data read for PQI
fs=25 #samplig frequency of S3
lowcut = 4
noFileRows =  2*durationWindow*fs
###############################################
#################################################
book_ppg = xlwt.Workbook(encoding="utf-8")

sheet_ppg = book_ppg.add_sheet("Sheet 2")

file_location = "achu4thjan_PPG.xlsx"
workbook = xlrd.open_workbook(file_location)
sheet = workbook.sheet_by_name('Sheet1')

ppgdata = []

book_ppg_featureSet = []
sheet_ppg_featureSet=[]

book_ppg_featureSet = xlwt.Workbook(encoding="utf-8")

#ppgDataColumns = range(1,4)
#for columnNumber in ppgDataColumns:
print('Reading the columns')

for rownum in range(sheet.nrows):
    ppgdata.append(sheet.cell_value(rownum, 3))
    # print('Value read:',sheet.cell(rownum,4))

print('Data length:', len(ppgdata))
n = len(ppgdata)
first10min=ppgdata[7500:22500]
last10min=ppgdata[n-15000:n]
print('EEG first:', len(first10min))
print('EEG last:', len(last10min))

ten_min_segments=[]
ten_min_segments=(first10min,last10min)
print('length of ten min seg:', len(ten_min_segments))
print('length of ten min seg:', len(ten_min_segments[0]))
print('length of ten min seg:', len(ten_min_segments[1]))
# Process data as many times as the number of fatigue flags

durationWindow = 10 # 10 seconds data read for PQI
fs=25 #samplig frequency of S3

noFileRows =  2*durationWindow*fs  # Number of rows to be read
#For example
# I want to read 10s data
# noFileRows = 2*10*25 #500
for x in range (0,len(ten_min_segments)):
    listDataChunks = []
    sf = 25  # sampling rate
    epochLength = 10;  # in Seconds
    overlap = 0
    time=sf*epochLength
    listDataChunks = split_overlap(ten_min_segments[x], sf * epochLength, overlap)
    print(len(listDataChunks))
    # Access one datachunk
    nRow_Position = 0
    if x == 0:
        fatigueFlag = 'N'
        healthstatus='normal'
    else:
        fatigueFlag = 'F'
        healthstatus='fatigue'

    #Create a worksheet where the data will be stored
    strWorksheetName = str("ppg")+str('_#')+str(fatigueFlag)
    print("Writing data into the worksheet:",strWorksheetName)
    sheet_ppg_featureSet = book_ppg_featureSet.add_sheet(strWorksheetName)


    sheet_ppg_featureSet.write(nRow_Position, 1, "mHR")
    sheet_ppg_featureSet.write(nRow_Position, 2, "stdHR")
    sheet_ppg_featureSet.write(nRow_Position, 3, "cvHR")
    sheet_ppg_featureSet.write(nRow_Position, 4, "RMSSD")
    sheet_ppg_featureSet.write(nRow_Position, 5, "pNN20")
    sheet_ppg_featureSet.write(nRow_Position, 6, "pNN30")
    sheet_ppg_featureSet.write(nRow_Position, 7, "pNN40")
    sheet_ppg_featureSet.write(nRow_Position, 8, "pNN50")
    sheet_ppg_featureSet.write(nRow_Position, 9, "LFPower")
    sheet_ppg_featureSet.write(nRow_Position, 10, "HFPower")
    sheet_ppg_featureSet.write(nRow_Position, 11, "TPower")
    sheet_ppg_featureSet.write(nRow_Position, 12, "LHRatio")
    sheet_ppg_featureSet.write(nRow_Position, 13, "class")

    nRow_Position= nRow_Position+1

    for itrn in range(len(listDataChunks)-1):
        print("\n \n The elements read for iteration ", itrn )
        # plt.plot(listDataChunks[itrn])
        # plt.show()
        # Define window length (4 seconds)
        if (type(listDataChunks[itrn][0] =='str')):
            listDataChunks[itrn].pop(0)

        data = listDataChunks[itrn]

        #if len(data)<12: sf =10

        if len(data) < (sf * epochLength)/2:
        #if len(data) < 4000:
            break; #Break out of the while loop and store all the data
        print('The number of signal points read is:',len(data))


        # Signal Conditioning

        #Identify trends and remove it
        smoothdata=savgol_filter(data,101,4)
        data_trendRemoved=data-smoothdata

        # Filtering (3)
        lowcut = 4
        fs=25
        filtereddata = butter_lowpass_filter(data_trendRemoved, lowcut, fs, order=3)


        # FFT Computation
        #Set variables

        f, Pxx=welch(filtereddata, fs=25.0, window=('gaussian',len(filtereddata)), nfft=len(filtereddata))

        peaksValIndex = peakutils.indexes(Pxx)
        identifiedPeaks=Pxx[peaksValIndex]
        MaxPeakVal_Index = max(enumerate(Pxx), key=lambda x: x[1])[0]

        xdatanew=[f[xyz] for xyz in peaksValIndex]

        hr_fft=max(xdatanew)*60


        #plt.semilogy(f, Pxx)
        #plt.ylim([0.5e-3, 1])
        #plt.ylabel('PSD [V**2/Hz]')
##        plt.subplot(2,1,2)
##        plt.xlabel('frequency [Hz]')
##        plt.ylabel('PSD')
##        plt.title('PSD')
##        plt.plot(f,Pxx)
##        plt.show()

        # PQI 1
        #if (len(peaksValIndex)<4 and (f[MaxPeakVal_Index] > 0.55 and f[MaxPeakVal_Index] < 2.5)): # Very restrictive
        if (len(peaksValIndex)<5 and (max(xdatanew) > 0.6 and max(xdatanew) < 2.5)): # 36 to 150 bpm
            print(" ")
        else:
            #color.write("Error 1: Unhealthy Segment. Data Ignored \n","COMMENT")
            #Comment the line below when you don't want the figures
            #PlotSignals()
            continue


        # Peakdetection in PPG Signal

        #Minimum distance between pulses = at least half-a-pulse width
        MIN_PULSE_DIST = math.ceil(fs/2)

        indexes = peakutils.indexes(filtereddata,thres=0.5,min_dist=MIN_PULSE_DIST)
        ydatanew=[filtereddata[xyz] for xyz in indexes]


        data=[]

        #Expected number of pulses based on heart rate
        noPulsesExpected = math.floor(max(xdatanew)*durationWindow)
        noActualPulses = len(indexes)

        #color.write("Info: Expected number of pulse is :  ","KEYWORD")
        print(noPulsesExpected)

        #color.write("Info: Actual number of pulse is :  ","KEYWORD")
        print(noActualPulses)

        from peakutils.plot import plot as pplot
        #plt.figure(figsize=(10,6))
        x=np.linspace(0,len(filtereddata)-1,len(filtereddata))

        plt.subplot(2,1,1)
        pplot(x, filtereddata, indexes)
        plt.xlabel('samples #')
        plt.ylabel('Amplitude')
        plt.title('Signal Pulses')

        #plt.title('First estimate')

        # PQI 2
        #if len(indexes)>10:
        if (noActualPulses < noPulsesExpected + 4) and (noActualPulses > noPulsesExpected - 4):
            print(" ")
        else:
            #color.write("Error 2: Unhealthy Segment. Data Ignored \n","COMMENT")
            #Comment the line below when you don't want the figures
            #PlotSignals()
            #plt.show()
            continue

        #color.write("Healthy Segment. Data is being Processed \n","STRING")
        #Comment the line below when you don't want the figures
        #PlotSignals()
        #plt.show()

        #"""
        # RR Interval calculation for 10 seconds data
        for i in range (0, len(indexes)-1):
                rr_interval_15secs.append(indexes[i+1] - indexes[i])

        BPM_inst = [(60*fs)/xy for xy in rr_interval_15secs]

        print('------------------------------------------')
        print('Measures calculated over 10-seconds')
        print('------------------------------------------')

        BPM_15secs = np.mean(BPM_inst);
        print("Mean of BPM (10 seconds):{:8.2f}  ".format(BPM_15secs))
        stdHR_15secs = np.std(BPM_inst)
        print("stdHR (10 seconds): {:8.2f} ".format(stdHR_15secs))

        rr_interval_15secs=[]

        BPM_inst=[]
        stdHR_15secs=[]
        #Add correction factor
        rr_peak_locs=indexes+lastpeakfound+correctionFactor
        lastpeakfound=rr_peak_locs[-1]


        """
        Calculation of correction factor -> Number of Samples that constitute a Half-pulse
        Madan, 8-November 2018
        """

        ######################################################

        # Total number of samples in the time-window  of consideration = time-window*sampling frequency
        # Number of samples in the time-window of consideration = noActualPulses
        # Number of samples contained in 1 pulse = time-window*sampling frequency/noActualPulses
        # Number of samples in half-pulse = 1/2*(time-window*sampling frequency/noActualPulses)

        # Half Pulse width
        ######################################################
        correctionFactor = math.ceil((durationWindow*fs)/(noActualPulses*2))

        #color.write("Info: Correction Factor is :  ","KEYWORD")
        print(correctionFactor)

        #The box which appends data as lists
        rr_Biglist.append(list(rr_peak_locs))

        #filtereddata=[]
        rr_peak_locs=[]

        """
        Code to check 18 bins i.e 3 Mins of aggregated data
        """
        durationAggregateData = 3 # 3 minutes of aggregated data
        LIST_FILL_SIZE = (durationAggregateData*60)/durationWindow

        if len(rr_Biglist)%LIST_FILL_SIZE ==0:

            #Process T.D., F.D., M.L.
            flat_list = np.hstack(rr_Biglist)
            print("flat list values",len(flat_list))
            # RR Interval calculation for 3 minutes data
            for i in range (0, len(flat_list)-1):
                rr_interval.append(flat_list[i+1] - flat_list[i])

            # OUTLIER DETECTION
            # Change in Outlier detection
            # Madan, 11-12-2019

            #USING HAMPEL
            rr_interval_hampel=hampel(rr_interval, k=5)
            iBPM = [(60*fs)/xy for xy in rr_interval_hampel]

            # removing > 2 sigma from mean
##            meanrr=np.mean(rr_interval,axis=0)
##            stdrr=np.std(rr_interval,axis=0)
##
##            rr_nooutlier=[x for x in rr_interval if (x<(meanrr+2*stdrr))]
##            rr_nooutlier=[x for x in rr_nooutlier if (x>(meanrr-2*stdrr))]
##
##            iBPM = [(60*fs)/xy for xy in rr_nooutlier]

            print('------------------------------------------')
            print('Time Domain Measures (3 minutes)')
            print('------------------------------------------')

            BPM = np.mean(iBPM);
            print("Mean of BPM: {:8.2f} ".format(BPM))
            stdHR = np.std(iBPM)
            #stdHR = np.around(stdHR,3)
            print("stdHR:  {:8.2f} ".format(stdHR))
            cvHR = BPM/stdHR
            cvHR = np.around(cvHR,3)
            print("cvHR:  {:8.2f} ".format(cvHR))

            # Madan, 11-12-2019
            # RR Interval in milliseconds
            rr_interval_ms=np.multiply(rr_interval_hampel,(1000/fs))
            #rr_interval_ms=np.multiply(rr_nooutlier,(1000/fs))

            rr_diff=[]
            for i in range (0, len(rr_interval_ms)-1):
                rr_diff.append(rr_interval_ms[i+1] - rr_interval_ms[i])

            rmssd =np.sqrt(np.mean(np.square(rr_diff)))


            print("RMSSD:  {:8.2f} ".format(rmssd))

            nn20 = [x for x in rr_diff if (abs(x)>20)]
            nn30 = [x for x in rr_diff if (abs(x)>30)]
            nn40 = [x for x in rr_diff if (abs(x)>40)]
            nn50 = [x for x in rr_diff if (abs(x)>50)]
            pnn20 = float(len(nn20)) / float(len(rr_diff))
            pnn30 = float(len(nn30)) / float(len(rr_diff))
            pnn40 = float(len(nn40)) / float(len(rr_diff))
            pnn50 = float(len(nn50)) / float(len(rr_diff))
            print("pNN20:  {:8.2f} ".format(pnn20))
            print("pNN30:  {:8.2f} ".format(pnn30))
            print("pNN40:  {:8.2f} ".format(pnn40))
            print("pNN50:  {:8.2f} ".format(pnn50))

            RR_x = flat_list[1:] #Remove the first entry, because first interval is assigned to the second beat.
            RR_y = iBPM #Y-values are equal to interval lengths

            # Getting equi-spaced HR samples through interpolation
            # Madan, 11-12-2019

            #f_hr_s=4 # 4 Hz sampling rate
            ##Create evenly spaced timeline starting at the second peak, its endpoint and length equal to position of last peak
            #RR_x_new = np.linspace(RR_x[0],RR_x[-1],1000/f_hr_s)
            #f = interp(RR_x_new, RR_x,RR_y) #Interpolate the signal with cubic spline interpolation

            # 11-12-2019
            Over_samp_fac=10 # 4 Hz sampling rate
            #Create evenly spaced timeline starting at the second peak, its endpoint and length equal to position of last peak
            RR_x_new = np.linspace(RR_x[0],RR_x[-1],np.size(RR_x)*Over_samp_fac)

            print(len(RR_x))
            print(len(RR_y))

            interp_func= interpolate.interp1d(RR_x,RR_y,kind='linear')

            RR_y_new=interp_func(RR_x_new)

            freq_s=int(Over_samp_fac*np.mean(BPM)/60)

            print('------------------------------------------')
            print('Frequency Domain Measures (3 minutes)')
            print('------------------------------------------')

            # welch
            # Madan, 11-12-2019
            #frq, Pxx1=welch(f, fs=5.0) #Sampling rate fixed at 5 Hz
            frq, Pxx1=welch(RR_y_new,freq_s) # Sampling rate set dynamically

            lf = np.trapz(Pxx1[(frq>=0.04) & (frq<=0.15)])
            #Slice frequency spectrum where x is between 0.04 and 0.15Hz (LF), and use NumPy's trapezoidal integration function to find the area
            print ("LF: {:8.2f}".format(lf))

            hf = np.trapz(Pxx1[(frq>=0.15) & (frq<=0.4)]) #Do the same for 0.16-0.5Hz (HF)
            print ("HF: {:8.2f}".format(hf))

            tp= np.trapz(Pxx1[(frq>=0.04) & (frq<=0.4)])#Do the same for 0.16-0.5Hz (HF)
            print ("TP:  {:8.2f}".format(tp))

            lhratio=lf/hf
            print ("LHRatio:  {:8.2f}".format(lhratio))

            sheet_ppg_featureSet.write(nRow_Position, 1, BPM)
            sheet_ppg_featureSet.write(nRow_Position, 2, stdHR)
            sheet_ppg_featureSet.write(nRow_Position, 3, cvHR)
            sheet_ppg_featureSet.write(nRow_Position, 4, rmssd)
            sheet_ppg_featureSet.write(nRow_Position, 5, pnn20)
            sheet_ppg_featureSet.write(nRow_Position, 6, pnn30)
            sheet_ppg_featureSet.write(nRow_Position, 7, pnn40)
            sheet_ppg_featureSet.write(nRow_Position, 8, pnn50)
            sheet_ppg_featureSet.write(nRow_Position, 9, lf)
            sheet_ppg_featureSet.write(nRow_Position, 10, hf)
            sheet_ppg_featureSet.write(nRow_Position, 11, tp)
            sheet_ppg_featureSet.write(nRow_Position, 12, lhratio)
            sheet_ppg_featureSet.write(nRow_Position, 13, healthstatus)
            nRow_Position= nRow_Position+1


            """
            Housekeeping
            ------------
            Popping off the first bin
            """
            popped=rr_Biglist.pop(0)
            rr_interval=[]
            flat_list=[]
            iBPM=[]
            #plt.show()
            #book.save('watch2output.xls')
##################################################################
            ###################################################
print("Writing into excel file")
book_ppg_featureSet.save('PPG_FeatureSet_approach1_'+str(CANDIDATE_NAME)+'.xls')
