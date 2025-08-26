from flask import Flask, render_template, request, send_file
import datetime as dt
import os
from osgeo import ogr, gdal
import geopandas as gpd

from version import __version__
import arcgis

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
    retval['projects'] = df_bap[['GlobalID_main', 'asset_name_main']].to_numpy()
        
    if(request.method.lower() == 'post'):
        if('btnGenerateBapFiles' in request.form):
            bap_project_id = request.form.get('ddlBapProject')
            
            bap_xl_report = r'C:\_kevin\test\bap_report.xlsx'
            bap_xl_component = r'C:\_kevin\test\bap_component.xlsx'
            bap_word_doc = r'C:\_kevin\test\bap.docx'
            arcgis.generate_bap_excel(df_bap, bap_project_id, bap_xl_report, bap_xl_component)
            arcgis.generate_bap_worddoc(df_bap, bap_project_id, arcgis.WORD_BAP_TEMPLATE, bap_word_doc)

            filebytes = arcgis.combine_bap_files(bap_xl_report, bap_xl_component, bap_word_doc)
            return send_file(filebytes, mimetype='application/zip', download_name='BAP.zip', as_attachment=True)

    return render_template(
        'index.jinja2', 
        **retval)

if __name__ == '__main__':
    app.run(debug=True)