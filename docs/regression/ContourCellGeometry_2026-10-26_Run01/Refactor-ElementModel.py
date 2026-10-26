from pathlib import Path
import re
p=Path('src/Section/CSectionModel.cls')
s=p.read_text(encoding='utf-8-sig')
if 'Private mConcreteElements()' in s: raise RuntimeError('Element model already refactored')
start=s.index('Public Sub AddConcreteElement(')
end=s.index('\nEnd Sub',start)+len('\nEnd Sub')
old=s[start:end]
signature=old[:old.index('    If area <= 0#')]
body='''    If area <= 0# Then Err.Raise vbObjectError + 5200, "CSectionModel", "Площадь бетонного элемента должна быть положительной."
    If Abs(rotation) > 2# * GEOM_PI Then rotation = NormalizeAnglePi(rotation)
    Dim explicitRectangle As Boolean
    explicitRectangle = StrComp(shapeType, "Rectangle", vbTextCompare) = 0 And width > GEOM_TOLERANCE And height > GEOM_TOLERANCE
    If explicitRectangle Then explicitRectangle = Abs(width * height - area) / area <= LOCAL_INERTIA_RECTANGLE_AREA_REL_TOLERANCE
    Dim exactBoundary As Boolean
    exactBoundary = StrComp(shapeType, "Contour", vbTextCompare) = 0
    If exactBoundary Then
        If Not ExactPositiveInertia(localIx, localIy, localIxy) Then Err.Raise vbObjectError + 5223, "CSectionModel", "Собственные инерции точного бетонного элемента должны задавать положительный тензор."
    ElseIf Not IsPositiveDefiniteInertia(localIx, localIy, localIxy) Then
        If explicitRectangle Then RectangleLocalInertiaComponents area, width, height, rotation, localIx, localIy, localIxy
    End If
    Dim displayWidth As Double, displayHeight As Double, displayRotation As Double, statusText As String
    displayWidth = width: displayHeight = height: displayRotation = rotation
    If Not exactBoundary And Not explicitRectangle Then
        Dim viewIx As Double, viewIy As Double, viewIxy As Double
        viewIx = localIx: viewIy = localIy: viewIxy = localIxy
        If Not IsPositiveDefiniteInertia(viewIx, viewIy, viewIxy) Then EquivalentSquareLocalInertiaComponents area, viewIx, viewIy, viewIxy
        If LocalInertiaIsNearlySquare(viewIx, viewIy, viewIxy) Then
            displayWidth = Sqr(area): displayHeight = displayWidth
        ElseIf Not TryInterpretAnisotropicBoundary(area, viewIx, viewIy, viewIxy, statusText, displayWidth, displayHeight, displayRotation) Then
            displayWidth = Sqr(area): displayHeight = displayWidth
        End If
    End If
    Dim nextIndex As Long, element As CConcreteElement
    nextIndex = mConcreteCount + 1
    Set element = New CConcreteElement
    element.Initialize "C" & CStr(nextIndex), xCoord, yCoord, area, materialID, sourceName, sourceHandle, shapeType, _
        displayWidth, displayHeight, displayRotation, rotationKnown Or Abs(rotation) > GEOM_TOLERANCE Or explicitRectangle, comment, localIx, localIy, localIxy
    EnsureConcreteCapacity nextIndex
    Set mConcreteElements(nextIndex) = element
    mConcreteX(nextIndex) = element.X: mConcreteY(nextIndex) = element.Y: mConcreteArea(nextIndex) = element.Area ' Неизменяемая числовая проекция для горячих циклов; источник данных — элемент.
    mConcreteCount = nextIndex
    mRevision = mRevision + 1
End Sub'''
s=s[:start]+signature+body+s[end:]
# One canonical array of immutable entities. Only X/Y/A retain a numerical projection.
mapping={'mConcreteID':'ElementID','mConcreteSourceName':'SourceName','mConcreteSourceHandle':'SourceHandle','mConcreteMaterialID':'MaterialID','mConcreteShapeType':'ShapeType','mConcreteWidth':'Width','mConcreteHeight':'Height','mConcreteRotation':'Rotation','mConcreteRotationKnown':'RotationKnown','mConcreteLocalIx':'LocalIx','mConcreteLocalIy':'LocalIy','mConcreteLocalIxy':'LocalIxy','mConcreteComment':'Comment'}
for field in list(mapping)+['mConcreteGeometry']:
    s=re.sub(r'^Private '+field+r'\(\).*\n','',s,flags=re.M)
    s=re.sub(r'^    Erase '+field+r'\n','',s,flags=re.M)
    s=re.sub(r'^    ReDim Preserve '+field+r'\(1 To capacity\)\n','',s,flags=re.M)
s=s.replace('Private mConcreteID() As String', 'Private mConcreteID() As String')
s=s.replace('Private mConcreteX() As Double', 'Private mConcreteElements() As CConcreteElement \' Канонические неизменяемые сущности; расчётные массивы ниже являются только числовой проекцией.\nPrivate mConcreteX() As Double',1)
s=s.replace('    Erase mConcreteX\n','    Erase mConcreteElements\n    Erase mConcreteX\n',1)
s=s.replace('    ReDim Preserve mConcreteX(1 To capacity)\n','    ReDim Preserve mConcreteElements(1 To capacity)\n    ReDim Preserve mConcreteX(1 To capacity)\n',1)
# Replace geometry publication atomically with a new immutable element.
s=s.replace('    Set mConcreteGeometry(index) = region.Clone\n    mConcreteShapeType(index) = "Contour"\n    mConcreteWidth(index) = widthValue: mConcreteHeight(index) = heightValue: mConcreteRotation(index) = rotationValue\n    mConcreteRotationKnown(index) = True', '    Set mConcreteElements(index) = mConcreteElements(index).WithGeometry(region, widthValue, heightValue, rotationValue)')
s=s.replace('    ConcreteHasBoundary = Not mConcreteGeometry(index) Is Nothing','    ConcreteHasBoundary = mConcreteElements(index).HasExactGeometry')
s=s.replace('Not mConcreteGeometry(index) Is Nothing','mConcreteElements(index).HasExactGeometry')
s=s.replace('mConcreteGeometry(index) Is Nothing','Not mConcreteElements(index).HasExactGeometry')
s=s.replace('mConcreteGeometry(index)','mConcreteElements(index).Geometry')
# Average legacy orientation must also create a new immutable entity.
s=s.replace('                mConcreteRotation(i) = fallbackRotation','                Set mConcreteElements(i) = mConcreteElements(i).WithRotation(fallbackRotation)')
for field,prop in mapping.items():
    s=re.sub(r'\b'+field+r'\(([^()]+)\)',lambda m:'mConcreteElements('+m[1]+').'+prop,s)
# Actual contours do not use the dimensional legacy fallback threshold.
needle='    If IsPositiveDefiniteInertia(localIx, localIy, localIxy) Then Exit Sub\n'
s=s.replace(needle,'    If mConcreteElements(index).HasExactGeometry Then\n        If Not ExactPositiveInertia(localIx, localIy, localIxy) Then Err.Raise vbObjectError + 5223, "CSectionModel", "Повреждены собственные инерции фактической бетонной ячейки."\n        Exit Sub\n    End If\n'+needle,1)
anchor="' Добавляет одно бетонное волокно или импортированную бетонную область."
helper='''\n' Возвращает неизменяемую сущность бетонного элемента; изменения модели выполняются её владельцем.
Public Function ConcreteElement(ByVal index As Long) As CConcreteElement
    CheckConcreteIndex index
    Set ConcreteElement = mConcreteElements(index)
End Function

' Проверяет положительность точного тензора без размерного нижнего порога старой оболочки.
' Нормированное произведение избегает переполнения и сохраняет тонкие физические сегменты.
Private Function ExactPositiveInertia(ByVal localIx As Double, ByVal localIy As Double, ByVal localIxy As Double) As Boolean
    If localIx <= 0# Or localIy <= 0# Then Exit Function
    ExactPositiveInertia = Abs(localIxy / Sqr(localIx) / Sqr(localIy)) < 1#
End Function

'''
s=s.replace(anchor,helper+anchor,1)
p.write_text(s,encoding='utf-8',newline='\r\n')
print('Canonical entity array installed; numerical X/Y/A projection retained.')