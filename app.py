from flask import Flask, render_template, request
from werkzeug.utils import secure_filename
from flask import send_file
import pandas as pd
from zipfile import ZipFile
app = Flask(__name__,template_folder='./templates')


def transform_shipping_method(value):
    if isinstance(value, str):
        split = value.split(" ")
        zona = split[1]
        if zona in ["Norte", "Sur", "Oeste", "Este"]:
            return zona
        else:
            return float('nan')
    else:
        return float('nan')
    return value

def file_processing():
    df = pd.read_csv("inputs/orders_export.csv")
    df["Shipping Method"] = df["Shipping Method"].map(transform_shipping_method)
    df["Created at"] = pd.to_datetime(df["Created at"]).dt.date

    df = df[['Name', 'Shipping Name', 'Email', 'Subtotal', 'Shipping', 'Total', 'Discount Code', 'Discount Amount', 'Shipping Method', 'Lineitem quantity', 'Lineitem name', 'Lineitem price', 'Lineitem compare at price', 'Lineitem sku', 'Lineitem requires shipping', 'Billing Name', 'Billing Street',
            'Billing Phone']]

    orders=df.Name.unique()
    for order in orders:
        client_name=df[df["Name"]==order]["Shipping Name"].unique()[0]
        df.loc[df.Name==order, 'Shipping Name'] =client_name




    """ df = df[['Name', 'Shipping Name', 'Email', 'Subtotal', 'Shipping', 'Total', 'Discount Code', 'Discount Amount', 'Shipping Method', 'Lineitem quantity', 'Lineitem name', 'Lineitem price', 'Lineitem compare at price', 'Lineitem sku', 'Lineitem requires shipping', 'Billing Name', 'Billing Street',
            'Billing Address1', 'Billing Phone', 'Shipping Street',
            'Shipping Address1', 'Shipping Phone']] """


    conv = pd.read_excel("inputs/factores.xlsx")

    #print(df.shape)
    df = pd.merge(df, conv, left_on="Lineitem sku", right_on="SKU", how='left')
    df["Peso_Lb"].fillna(1, inplace=True)
    df["Peso_Kg"] = df["Peso_Lb"]/2.20462

    df["Total_Lb"] = df["Peso_Lb"]*df["Lineitem quantity"]

    df["Total_Kg"] = df["Total_Lb"]/2.20462

    df["Total Item"]=df["Lineitem quantity"]*df["Lineitem price"]
    #print(df)

    df.to_excel("ventas.xlsx")
    print("Archivo de ventas generado: ",df.shape[0]," lineas procesadas.")
    print("Generando consolidado de ventas...")
    consolidado = df[["SKU", "Producto", "Total_Lb", "Total_Kg","Lineitem quantity","Presentacion"]
                    ].groupby(["Producto", "SKU"]).sum()

    #print(consolidado)
    consolidado.to_excel("consolidado.xlsx")

@app.route('/')
def upload_file():
   return render_template('upload.html')
	
@app.route('/uploader', methods = ['GET', 'POST'])
def upload_filePost():
   if request.method == 'POST':
       try:
           f = request.files['orders']
           f.save("inputs/orders_export.csv")
           f = request.files['factors']
           f.save("inputs/factores.xlsx")
           file_processing()
           zipObj = ZipFile('resultados.zip', 'w')
           zipObj.write('ventas.xlsx')
           zipObj.write('consolidado.xlsx')
           zipObj.close()
           return send_file('resultados.zip',as_attachment=True)
       except Exception as e:
           return str(e)

      
      #return 'file uploaded successfully'
		
if __name__ == '__main__':
   app.run()