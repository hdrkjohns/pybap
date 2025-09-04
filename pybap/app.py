from flask import Flask, render_template, request, send_file
import datetime as dt
import os
import io
from osgeo import ogr, gdal
import geopandas as gpd

from .version import __version__
from . import arcgis

app  = Flask(__name__,
             template_folder='templates',
             static_folder='static')

@app.route('/', methods=['post', 'get'])
def index():
    retval = {}

    today = dt.datetime.today().strftime('%Y-%m-%d %H:%M')
    retval['today'] = today
    retval['version'] = __version__

    df_bap = arcgis.bap_gdb_to_dataframe(os.path.join(arcgis.OUT_DIR, 'BAP.gdb'))
    if(df_bap is not None):
        retval['projects'] = df_bap[['GlobalID_main', 'asset_name_main']].to_numpy()
    else:
        retval['projects'] = [(-1, 'Projects Could Not Be Loaded!')]

    if(request.method.lower() == 'post'):
        if('btnGenerateBapFiles' in request.form):
            bap_project_id = request.form.get('ddlBapProject')

            wb_report, wb_component = arcgis.generate_bap_excel(df_bap, bap_project_id)
            doc = arcgis.generate_bap_worddoc(df_bap, bap_project_id, arcgis.WORD_BAP_TEMPLATE)
            wb_report_buff = io.BytesIO()
            wb_component_buff = io.BytesIO()
            bap_doc_buff = io.BytesIO()

            wb_report.save(wb_report_buff)
            wb_component.save(wb_component_buff)
            doc.save(bap_doc_buff)

            filebytes = arcgis.combine_bap_files(wb_report_buff, wb_component_buff, bap_doc_buff)

            del wb_report_buff
            del wb_component_buff
            del bap_doc_buff

            return send_file(filebytes, mimetype='application/zip', download_name='BAP.zip', as_attachment=True)

    return render_template(
        'index.jinja2',
        **retval)

if(__name__ == '__main__'):
    app.run(debug=True)