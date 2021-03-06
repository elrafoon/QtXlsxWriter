/****************************************************************************
** Copyright (c) 2013-2014 Debao Zhang <hello@debao.me>
** All right reserved.
**
** Permission is hereby granted, free of charge, to any person obtaining
** a copy of this software and associated documentation files (the
** "Software"), to deal in the Software without restriction, including
** without limitation the rights to use, copy, modify, merge, publish,
** distribute, sublicense, and/or sell copies of the Software, and to
** permit persons to whom the Software is furnished to do so, subject to
** the following conditions:
**
** The above copyright notice and this permission notice shall be
** included in all copies or substantial portions of the Software.
**
** THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
** EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
** MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
** NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
** LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
** OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
** WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
**
****************************************************************************/

#ifndef QXLSX_CHART_H
#define QXLSX_CHART_H

#include <math.h>
#include "xlsxabstractooxmlfile.h"
#include "xlsxcolor_p.h"

#include <QSharedPointer>

class QXmlStreamReader;
class QXmlStreamWriter;

QT_BEGIN_NAMESPACE_XLSX

class AbstractSheet;
class Worksheet;
class ChartPrivate;
class CellRange;
class DrawingAnchor;

class Q_XLSX_EXPORT ChartFill {
public:
    enum FillStyle {
        FS_Solid = 1
    };

    ChartFill();
    ChartFill(const XlsxColor &color, FillStyle style = FS_Solid);

    bool operator !=(const ChartFill &other) const;

    XlsxColor color;
    FillStyle style;
};

class Q_XLSX_EXPORT ChartLine {
public:
    ChartLine();
    ChartLine(const ChartFill &fill, int width = 10000);

    bool operator !=(const ChartLine &other) const;

    ChartFill fill;
    int width;
};

class Q_XLSX_EXPORT ChartShape {
public:
    ChartShape(ChartFill area = ChartFill(), ChartLine line = ChartLine());

    bool operator !=(const ChartShape &other) const;

    ChartFill area;
    ChartLine line;
};

class Q_XLSX_EXPORT XlsxAxis
{
public:
    enum Type
    {
        T_Cat,
        T_Val,
        T_Date,
        T_Ser
    };

    enum Pos
    {
        Left,
        Right,
        Top,
        Bottom
    };

    XlsxAxis() : min(nanf("")), max(nanf("")) {}

    XlsxAxis(Type t, Pos p, int id, int crossId, QString title = "", double min = nanf(""), double max = nanf(""))
        :type(t), axisPos(p), axisId(id), crossAx(crossId), title(title), min(min), max(max)
    {
    }

    Type type;
    Pos axisPos; //l,r,b,t
    int axisId;
    int crossAx;
    QString title;
    double min, max;
};



class Q_XLSX_EXPORT Chart : public AbstractOOXmlFile
{
    Q_DECLARE_PRIVATE(Chart)

public:
    enum ChartType {
        CT_Area = 1, //Zero is internally used for unknown types
        CT_Area3D,
        CT_Line,
        CT_Line3D,
        CT_Stock,
        CT_Radar,
        CT_Scatter,
        CT_Pie,
        CT_Pie3D,
        CT_Doughnut,
        CT_Bar,
        CT_Bar3D,
        CT_OfPie,
        CT_Surface,
        CT_Surface3D,
        CT_Bubble
    };

    enum ChartStyle {
        CS_Default = 0,
        CS_Line
    };

    enum MarkerType {
        MT_DEFAULT = 0,
        MT_NONE
    };

    ~Chart();

    void addSeries(const CellRange &range, AbstractSheet *sheet=0, MarkerType marker = MT_DEFAULT,
                   QSharedPointer<ChartShape> shape = 0);
    void addXYSeries(const CellRange &x, const CellRange &y, AbstractSheet *sheet=0, MarkerType marker = MT_DEFAULT,
                   QSharedPointer<ChartShape> shape = 0);
    void setChartType(ChartType type);
    void setChartStyle(ChartStyle style);
    // use only when overriding default axis
    void addAxis(QSharedPointer<XlsxAxis> axis);
    void clearAxis();
    unsigned axisCount() const;
    QSharedPointer<XlsxAxis> axis(unsigned index) const;

    void saveToXmlFile(QIODevice *device) const;
    bool loadFromXmlFile(QIODevice *device);

private:
    friend class AbstractSheet;
    friend class Worksheet;
    friend class Chartsheet;
    friend class DrawingAnchor;

    Chart(AbstractSheet *parent, CreateFlag flag);
};

QT_END_NAMESPACE_XLSX

#endif // QXLSX_CHART_H

