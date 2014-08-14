PoiCreateExcel
==============

使用POI產生Excel(同時支援.xls與.xlsx)

==============
POI是Java的一個處理Excel的套件
POI對於.xls與.xlsx使用的實作類別是不同的
網路上範例很多是使用HSSF開頭的，也就是專門處理.xls

但POI本身的設計其實可以讓開發者依賴於interface進行開發
也就是在一開始建立WorkBook的實體時，用類似簡單工廠模式的方式使用對應的實作類別，就能產生.xls或.xlsx
在處理邏輯上可以兩者通用而不用寫兩個版本

兩種實作類別需要使用的jar檔也是不太一樣的
但因為此程式有使用gradle管理相依性，所以不用特別找相依的jar檔
