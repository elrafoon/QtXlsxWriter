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

#ifndef QXLSX_CHART_P_H
#define QXLSX_CHART_P_H

//
//  W A R N I N G
//  -------------
//
// This file is not part of the Qt Xlsx API.  It exists for the convenience
// of the Qt Xlsx.  This header file may change from
// version to version without notice, or even be removed.
//
// We mean it.
//

#include "xlsxabstractooxmlfile_p.h"
#include "xlsxchart.h"

#include <QSharedPointer>

class QXmlStreamReader;
class QXmlStreamWriter;

namespace QXlsx {

class XlsxSeries
{
public:
    //At present, we care about number cell ranges only!
    QString numberDataSource_numRef; //yval, val
    QString axDataSource_numRef; //xval, cat
    Chart::MarkerType markerType;
    QSharedPointer<ChartShape> shape;
};

class ChartPrivate : public AbstractOOXmlFilePrivate
{
    Q_DECLARE_PUBLIC(Chart)

public:
    ChartPrivate(Chart *q, Chart::CreateFlag flag);
    ~ChartPrivate();

    bool loadXmlChart(QXmlStreamReader &reader);
    bool loadXmlPlotArea(QXmlStreamReader &reader);
    bool loadXmlXxxChart(QXmlStreamReader &reader);
    bool loadXmlSer(QXmlStreamReader &reader);
    bool loadXmlShape(QXmlStreamReader &reader, ChartShape &shape);
    bool loadXmlLine(QXmlStreamReader &reader, ChartLine &line);
    bool loadXmlFill(QXmlStreamReader &reader, ChartFill &fill);
    QString loadXmlNumRef(QXmlStreamReader &reader);
    bool loadXmlAxis(QXmlStreamReader &reader);
    bool loadXmlAxisTitle(QXmlStreamReader &reader, XlsxAxis &axis);

    void saveXmlChart(QXmlStreamWriter &writer) const;
    void saveXmlPieChart(QXmlStreamWriter &writer) const;
    void saveXmlBarChart(QXmlStreamWriter &writer) const;
    void saveXmlLineChart(QXmlStreamWriter &writer) const;
    void saveXmlScatterChart(QXmlStreamWriter &writer) const;
    void saveXmlAreaChart(QXmlStreamWriter &writer) const;
    void saveXmlDoughnutChart(QXmlStreamWriter &writer) const;
    void saveXmlSer(QXmlStreamWriter &writer, XlsxSeries *ser, int id) const;
    void saveXmlShape(QXmlStreamWriter &writer, const ChartShape &shape) const;
    void saveXmlFill(QXmlStreamWriter &writer, const ChartFill &fill) const;
    void saveXmlLine(QXmlStreamWriter &writer, const ChartLine &line) const;
    void saveXmlAxes(QXmlStreamWriter &writer) const;
    void saveXmlAxisTitle(QXmlStreamReader &reader) const;

    Chart::ChartType chartType;
    Chart::ChartStyle chartStyle;

    QList<QSharedPointer<XlsxSeries> > seriesList;
    QList<QSharedPointer<XlsxAxis> > axisList;

    AbstractSheet *sheet;
};

} // namespace QXlsx

#endif // QXLSX_CHART_P_H
