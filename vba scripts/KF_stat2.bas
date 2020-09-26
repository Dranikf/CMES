Attribute VB_Name = "Module1"
Function Log10(X)
'почемуто не получается использовать встроенный Log10
'на формуах посоветовали такой выход
    Log10 = Log(X) / Log(10#)
End Function

Function k_for_sterjers(arr)
'функция вычисляет значение k (число в знаменателе)
'для формулы стержерса
'arr - массив для которого ведуться вычисления
    k_for_sterjers = 1 + 3.3221 * Log10(arr.Count)
End Function

Function sterjers_formula(arr)
    'функция для вычиления длинны интервалов по формуле стержерса
    Max = Application.WorksheetFunction.Max(arr)
    Min = Application.WorksheetFunction.Min(arr)
    
    k = k_for_sterjers(arr)
    
    sterjers_formula = (Max - Min) / k
    
End Function

Function interval_mediana(low_inter, inter_size, sum_freq, freq, n)
    'фукнция для расчета медианы для интевального ряда
    'low_inter - нижняя граница интервала
    'inter_size - размер интевала
    'sum_freq - сумма накопленных чатот до интервала
    'freq - медианного интервала
    'n - объем выборки
    interval_mediana = low_inter + inter_size * ((n / 2) - sum_freq) / freq
End Function
