# -*- coding: utf-8 -*-
"""
ECG Detection Engine (EDE) - Golden Reference Algorithm
2022 IC Design Contest (Graduate Level) / NCHU CAD/VLSI Final Project
"""

import numpy as np

def moving_average_16taps(x):
    """
    Step 1: 16-tap Moving Average
    MV(n) = [x(n) + x(n-1) + ... + x(n-15)] / 16
    Out-of-bound samples (k < 0) are replaced by x(0).
    """
    N = len(x)
    MV = np.zeros(N, dtype=np.float64)
    for n in range(N):
        s = 0.0
        for k in range(16):
            idx = n - k
            if idx < 0:
                s += x[0]
            else:
                s += x[idx]
        MV[n] = s / 16.0
    return MV

def remove_dc_bias(x, MV):
    """
    Step 2: Remove DC-Bias
    Del_DC(n) = x(n) - MV(n)
    """
    return x - MV

def windowing_filter_15taps(Del_DC):
    """
    Step 3: 15-tap Windowing Filter
    WF(n) = sum_{k=-7}^{7} [Del_DC(n+k) * C(k+7)] / 64
    C[0:14] = [1, 2, 3, 4, 5, 6, 7, 8, 7, 6, 5, 4, 3, 2, 1]
    Out-of-bound left (< 0) -> Del_DC(0)
    Out-of-bound right (>= N) -> Del_DC(N-1)
    """
    N = len(Del_DC)
    C = np.array([1, 2, 3, 4, 5, 6, 7, 8, 7, 6, 5, 4, 3, 2, 1], dtype=np.float64)
    WF = np.zeros(N, dtype=np.float64)
    
    for n in range(N):
        s = 0.0
        for k in range(-7, 8):
            idx = n + k
            c_val = C[k + 7]
            if idx < 0:
                val = Del_DC[0]
            elif idx >= N:
                val = Del_DC[N - 1]
            else:
                val = Del_DC[idx]
            s += val * c_val
        WF[n] = s / 64.0
    return WF

def detect_qrs(WF, samples_per_beat=120):
    """
    Step 4: Q, R, S Points Detection
    Total: 2400 samples / 120 = 20 Beats.
    For each Beat:
      1. R point: Global Max in this 120-sample window.
      2. Q point: Search backward from R for the first Local Min.
      3. S point: Search forward from R for the first Local Min.
      Tie-breaking rule: If WF values are equal, pick the one closest to R.
    """
    N = len(WF)
    num_beats = N // samples_per_beat
    qrs_results = []
    
    for b in range(num_beats):
        start_idx = b * samples_per_beat
        end_idx = start_idx + samples_per_beat  # exclusive
        beat_wf = WF[start_idx:end_idx]
        
        # 1. Find Global Max (R point)
        # Note: In case of duplicate global max, standard convention usually takes the first or highest prominence.
        r_rel = int(np.argmax(beat_wf))
        r_abs = start_idx + r_rel
        
        # 2. Search backward from R for Q point (Local Min)
        # Scan backward from r_abs - 1 towards start_idx
        q_cand = r_abs
        for curr in range(r_abs - 1, start_idx - 1, -1):
            if curr == start_idx:
                # Reached left boundary of the beat
                q_cand = curr
                break
            # Check local minimum: WF(curr) <= WF(curr+1) and WF(curr) <= WF(curr-1)
            # When scanning from R to left, values decrease then increase.
            # The turnaround occurs when WF(curr - 1) > WF(curr).
            if WF[curr - 1] > WF[curr]:
                q_cand = curr
                break
        q_abs = q_cand
        
        # 3. Search forward from R for S point (Local Min)
        # Scan forward from r_abs + 1 towards end_idx - 1
        s_cand = r_abs
        for curr in range(r_abs + 1, end_idx):
            if curr == end_idx - 1:
                # Reached right boundary of the beat
                s_cand = curr
                break
            # When scanning from R to right, values decrease then increase.
            # The turnaround occurs when WF(curr + 1) > WF(curr).
            if WF[curr + 1] > WF[curr]:
                s_cand = curr
                break
        s_abs = s_cand
        
        qrs_results.append((b, q_abs, r_abs, s_abs))
        
    return qrs_results

def run_pipeline(x):
    """Executes the full EDE pipeline on input samples x."""
    MV = moving_average_16taps(x)
    Del_DC = remove_dc_bias(x, MV)
    WF = windowing_filter_15taps(Del_DC)
    qrs = detect_qrs(WF)
    return MV, Del_DC, WF, qrs

if __name__ == "__main__":
    print("EDE Golden Pipeline module loaded successfully.")

