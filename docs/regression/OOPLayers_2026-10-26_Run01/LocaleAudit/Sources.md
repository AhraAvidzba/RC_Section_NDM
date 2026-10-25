# Источники для локального теста локалей

Microsoft VBA [Type conversion functions](https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/type-conversion-functions) описывает зависимость строковых преобразований от локали и признание разделителей тысяч. Поэтому механическая замена CDbl не является решением.

Windows [NLS terminology](https://learn.microsoft.com/en-us/windows/win32/intl/nls-terminology) и [SetThreadLocale](https://learn.microsoft.com/en-us/windows/win32/api/winnls/nf-winnls-setthreadlocale) описывают локаль вызывающего потока; её изменение не гарантирует влияния на любой API. В нативном тесте отдельно записываются GetThreadLocale и фактический CStr/Format, чтобы не приравнивать сам успешный вызов SetThreadLocale к изменению VBA-локали. User locale глобально не меняется.
