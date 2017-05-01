cskyHTML = """
<html>
    <head></head>
    <body>
        <table style="border-collapse: collapse; width: 45%; display: inline-table">
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Month To Date</th>
                <th style="border: 1px solid black">Total</th>
                <th style="border: 1px solid black">LP14</th>
                <th style="border: 1px solid black">LP21</th>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Sales</th>
                <td style="border: 1px solid black" align="center">{All.salesTotal}</td>
                <td style="border: 1px solid black" align="center">{LP14.salesTotal}</td>
                <td style="border: 1px solid black" align="center">{LP21.salesTotal}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Projection</th>
                <td style="border: 1px solid black" align="center">{All.salesProject}</td>
                <td style="border: 1px solid black" align="center">{LP14.salesProject}</td>
                <td style="border: 1px solid black" align="center">{LP21.salesProject}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Plan</th>
                <td style="border: 1px solid black" align="center">{All.salesGoal}</td>
                <td style="border: 1px solid black" align="center">{LP14.salesGoal}</td>
                <td style="border: 1px solid black" align="center">{LP21.salesGoal}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Percent to Plan</th>
                <td style="border: 1px solid black" align="center">{All.salesPercent}%</td>
                <td style="border: 1px solid black" align="center">{LP14.salesPercent}%</td>
                <td style="border: 1px solid black" align="center">{LP21.salesPercent}%</td>
            </tr>
        </table>
        <table style="border-collapse: collapse; width: 45%; display: inline-table">
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Quarter To Date</th>
                <th style="border: 1px solid black">Total</th>
                <th style="border: 1px solid black">LP14</th>
                <th style="border: 1px solid black">LP21</th>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Sales</th>
                <td style="border: 1px solid black" align="center">{All.salesTotalQ}</td>
                <td style="border: 1px solid black" align="center">{LP14.salesTotalQ}</td>
                <td style="border: 1px solid black" align="center">{LP21.salesTotalQ}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Projection</th>
                <td style="border: 1px solid black" align="center">{All.salesProjectQ}</td>
                <td style="border: 1px solid black" align="center">{LP14.salesProjectQ}</td>
                <td style="border: 1px solid black" align="center">{LP21.salesProjectQ}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Plan</th>
                <td style="border: 1px solid black" align="center">{All.salesGoalQ}</td>
                <td style="border: 1px solid black" align="center">{LP14.salesGoalQ}</td>
                <td style="border: 1px solid black" align="center">{LP21.salesGoalQ}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Percent to Plan</th>
                <td style="border: 1px solid black" align="center">{All.salesPercentQ}%</td>
                <td style="border: 1px solid black" align="center">{LP14.salesPercentQ}%</td>
                <td style="border: 1px solid black" align="center">{LP21.salesPercentQ}%</td>
            </tr>
        </table>
        <p></p>
        <table style="border-collapse: collapse; width: 45%; display: inline-table">
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Month To Date</th>
                <th style="border: 1px solid black">Total</th>
                <th style="border: 1px solid black">LP14</th>
                <th style="border: 1px solid black">LP21</th>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Car Count</th>
                <td style="border: 1px solid black" align="center">{All.carsTotal}</td>
                <td style="border: 1px solid black" align="center">{LP14.carsTotal}</td>
                <td style="border: 1px solid black" align="center">{LP21.carsTotal}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Projection</th>
                <td style="border: 1px solid black" align="center">{All.carsProject}</td>
                <td style="border: 1px solid black" align="center">{LP14.carsProject}</td>
                <td style="border: 1px solid black" align="center">{LP21.carsProject}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Plan</th>
                <td style="border: 1px solid black" align="center">{All.carsGoal}</td>
                <td style="border: 1px solid black" align="center">{LP14.carsGoal}</td>
                <td style="border: 1px solid black" align="center">{LP21.carsGoal}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Percent to Plan</th>
                <td style="border: 1px solid black" align="center">{All.carsPercent}%</td>
                <td style="border: 1px solid black" align="center">{LP14.carsPercent}%</td>
                <td style="border: 1px solid black" align="center">{LP21.carsPercent}%</td>
            </tr>
        </table>
        <table style="border-collapse: collapse; width: 45%; display: inline-table">
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Quarter To Date</th>
                <th style="border: 1px solid black">Total</th>
                <th style="border: 1px solid black">LP14</th>
                <th style="border: 1px solid black">LP21</th>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Car Count</th>
                <td style="border: 1px solid black" align="center">{All.carsTotalQ}</td>
                <td style="border: 1px solid black" align="center">{LP14.carsTotalQ}</td>
                <td style="border: 1px solid black" align="center">{LP21.carsTotalQ}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Projection</th>
                <td style="border: 1px solid black" align="center">{All.carsProjectQ}</td>
                <td style="border: 1px solid black" align="center">{LP14.carsProjectQ}</td>
                <td style="border: 1px solid black" align="center">{LP21.carsProjectQ}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Plan</th>
                <td style="border: 1px solid black" align="center">{All.carsGoalQ}</td>
                <td style="border: 1px solid black" align="center">{LP14.carsGoalQ}</td>
                <td style="border: 1px solid black" align="center">{LP21.carsGoalQ}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Percent to Plan</th>
                <td style="border: 1px solid black" align="center">{All.carsPercentQ}%</td>
                <td style="border: 1px solid black" align="center">{LP14.carsPercentQ}%</td>
                <td style="border: 1px solid black" align="center">{LP21.carsPercentQ}%</td>
            </tr>
        </table>
        <p></p>
        <table style="border-collapse: collapse; width: 45%; display: inline-table">
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Month To Date</th>
                <th style="border: 1px solid black">Total</th>
                <th style="border: 1px solid black">LP14</th>
                <th style="border: 1px solid black">LP21</th>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Labor</th>
                <td style="border: 1px solid black" align="center">{All.laborTotal}</td>
                <td style="border: 1px solid black" align="center">{LP14.laborTotal}</td>
                <td style="border: 1px solid black" align="center">{LP21.laborTotal}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Projection</th>
                <td style="border: 1px solid black" align="center">{All.laborProject}</td>
                <td style="border: 1px solid black" align="center">{LP14.laborProject}</td>
                <td style="border: 1px solid black" align="center">{LP21.laborProject}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Plan</th>
                <td style="border: 1px solid black" align="center">{All.laborGoal}</td>
                <td style="border: 1px solid black" align="center">{LP14.laborGoal}</td>
                <td style="border: 1px solid black" align="center">{LP21.laborGoal}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Percent to Plan</th>
                <td style="border: 1px solid black" align="center">{All.laborPercent}%</td>
                <td style="border: 1px solid black" align="center">{LP14.laborPercent}%</td>
                <td style="border: 1px solid black" align="center">{LP21.laborPercent}%</td>
            </tr>
        </table>
        <table style="border-collapse: collapse; width: 45%; display: inline-table">
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Quarter To Date</th>
                <th style="border: 1px solid black">Total</th>
                <th style="border: 1px solid black">LP14</th>
                <th style="border: 1px solid black">LP21</th>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Labor</th>
                <td style="border: 1px solid black" align="center">{All.laborTotalQ}</td>
                <td style="border: 1px solid black" align="center">{LP14.laborTotalQ}</td>
                <td style="border: 1px solid black" align="center">{LP21.laborTotalQ}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Projection</th>
                <td style="border: 1px solid black" align="center">{All.laborProjectQ}</td>
                <td style="border: 1px solid black" align="center">{LP14.laborProjectQ}</td>
                <td style="border: 1px solid black" align="center">{LP21.laborProjectQ}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Plan</th>
                <td style="border: 1px solid black" align="center">{All.laborGoalQ}</td>
                <td style="border: 1px solid black" align="center">{LP14.laborGoalQ}</td>
                <td style="border: 1px solid black" align="center">{LP21.laborGoalQ}</td>
            </tr>
            <tr style="border: 1px solid black">
                <th style="border: 1px solid black">Percent to Plan</th>
                <td style="border: 1px solid black" align="center">{All.laborPercentQ}%</td>
                <td style="border: 1px solid black" align="center">{LP14.laborPercentQ}%</td>
                <td style="border: 1px solid black" align="center">{LP21.laborPercentQ}%</td>
            </tr>
        </table>
    </body>
</html>
    """