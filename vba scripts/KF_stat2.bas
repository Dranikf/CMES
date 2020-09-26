Attribute VB_Name = "Module1"
Function Log10(X)
'�������� �� ���������� ������������ ���������� Log10
'�� ������� ������������ ����� �����
    Log10 = Log(X) / Log(10#)
End Function

Function k_for_sterjers(arr)
'������� ��������� �������� k (����� � �����������)
'��� ������� ���������
'arr - ������ ��� �������� �������� ����������
    k_for_sterjers = 1 + 3.3221 * Log10(arr.Count)
End Function

Function sterjers_formula(arr)
    '������� ��� ��������� ������ ���������� �� ������� ���������
    Max = Application.WorksheetFunction.Max(arr)
    Min = Application.WorksheetFunction.Min(arr)
    
    k = k_for_sterjers(arr)
    
    sterjers_formula = (Max - Min) / k
    
End Function

Function interval_mediana(low_inter, inter_size, sum_freq, freq, n)
    '������� ��� ������� ������� ��� ������������ ����
    'low_inter - ������ ������� ���������
    'inter_size - ������ ��������
    'sum_freq - ����� ����������� ����� �� ���������
    'freq - ���������� ���������
    'n - ����� �������
    interval_mediana = low_inter + inter_size * ((n / 2) - sum_freq) / freq
End Function
