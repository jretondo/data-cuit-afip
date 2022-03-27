const XLSX = require("xlsx")
const path = require("path")
const Afip = require("ts-afip-ws")
const verificadorCuit = require('verificar-cuit-ar')

const getDataSheet = () => {
    const file = path.join(__dirname, "Files", "Input", "original.xlsx")
    const workBook = XLSX.readFile(file)
    const sheets = workBook.SheetNames;
    const hoja1 = sheets[0]
    return XLSX.utils.sheet_to_json(workBook.Sheets[hoja1])
}

const createAfipObj = () => {
    return new Afip({
        CUIT: 20350925148,
        res_folder: path.join(__dirname, "Afip", "crt"),
        cert: "203509251484.crt",
        key: "203509251484.key",
        ta_folder: path.join(__dirname, "Afip", "token"),
        production: true
    })
}

const getDataAfip = (afip, CUIT) => {
    return new Promise((resolve, reject) => {
        resolve(afip.RegisterScopeFive.getTaxpayerDetails(CUIT))
    })
}

const principal = async () => {
    const dataOriginal = getDataSheet()
    const Afip = createAfipObj()
    let rowsArray = []
    for (let i = 0; i < dataOriginal.length; i++) {
        let row = {}
        const cuitPuro = dataOriginal[i].CUIL
        const cuit = cuitPuro.replace(/[^0-9\.]+/g, '')
        const id = dataOriginal[i].CUENTA
        const nombre = dataOriginal[i].TITULAR
        const verificar = await verificadorCuit(cuit)
        if (verificar.isCuit) {
            const dataAfip = await getDataAfip(Afip, cuit)
            let tipoPersona = ""
            let error = false
            try {
                tipoPersona = dataAfip.datosGenerales.tipoPersona
            } catch (err) {
                error = true
            }
            if (tipoPersona === undefined) {
                error = true
            }
            if (!error) {
                let razSocial = ""
                if (tipoPersona === "FISICA") {
                    razSocial = (dataAfip.datosGenerales.apellido === undefined ? "" : dataAfip.datosGenerales.apellido) + " " + (dataAfip.datosGenerales.nombre === undefined ? "" : dataAfip.datosGenerales.nombre)
                } else {
                    razSocial = dataAfip.datosGenerales.razonSocial
                }
                let direccion = dataAfip.datosGenerales.domicilioFiscal.direccion + ", " + dataAfip.datosGenerales.domicilioFiscal.localidad + ", " + dataAfip.datosGenerales.domicilioFiscal.descripcionProvincia
                const datosMonotributo = dataAfip.datosMonotributo
                const datosRegGral = dataAfip.datosRegimenGeneral
                if (datosMonotributo === undefined) {
                    if (datosRegGral === undefined) {
                        row = {
                            CUENTA: id,
                            TITULAR: nombre,
                            CUIL: cuit,
                            "RAZON SOCIAL": razSocial,
                            "COND. IVA": "Error",
                            "ERROR": "No inscripto a ningún impuesto",
                            "DIRECCION": direccion,
                            "ACTIVIDAD": "",
                        }
                        rowsArray.push(row)
                    } else {
                        const impuestos = datosRegGral.impuesto
                        let actividad = ""
                        try {
                            actividad = datosRegGral.actividad[0].descripcionActividad
                        } catch (error) {

                        }
                        let impuestoStr = "ERROR EN LOS IMPUESTOS"
                        impuestos.map(item => {
                            const idImp = item.idImpuesto
                            if (idImp === 30) {
                                impuestoStr = "IVA INSCRIPTO"
                            } else if (idImp === 32) {
                                impuestoStr = "IVA EXENTO"
                            } else if (idImp === 33) {
                                impuestoStr = "IVA RESPONSABLE NO INSCRIPTO"
                            } else if (idImp === 34) {
                                impuestoStr = "IVA NO ALCANZADO"
                            }
                        })
                        row = {
                            CUENTA: id,
                            TITULAR: nombre,
                            CUIL: cuit,
                            "RAZON SOCIAL": razSocial,
                            "COND. IVA": impuestoStr,
                            "ERROR": "",
                            "DIRECCION": direccion,
                            "ACTIVIDAD": actividad,
                        }
                        rowsArray.push(row)
                    }
                } else {
                    const condiva = "MONOTRIBUTISTA " + datosMonotributo.categoriaMonotributo.descripcionCategoria
                    let actividad = ""
                    try {
                        actividad = datosMonotributo.actividadMonotributista.descripcionActividad
                    } catch (error) {

                    }
                    row = {
                        CUENTA: id,
                        TITULAR: nombre,
                        CUIL: cuit,
                        "RAZON SOCIAL": razSocial,
                        "COND. IVA": condiva,
                        "ERROR": "",
                        "DIRECCION": direccion,
                        "ACTIVIDAD": actividad,
                    }
                    rowsArray.push(row)
                }
            } else {
                if (dataAfip === null) {
                    row = {
                        CUENTA: id,
                        TITULAR: nombre,
                        CUIL: cuit,
                        "RAZON SOCIAL": "",
                        "COND. IVA": "Error",
                        "ERROR": "No hay datos en la consulta",
                        "DIRECCION": "",
                        "ACTIVIDAD": "",
                    }
                    rowsArray.push(row)
                } else {
                    const razSocial = dataAfip.errorConstancia.apellido + " " + dataAfip.errorConstancia.nombre
                    const errorStr = dataAfip.errorConstancia.error[0]
                    row = {
                        CUENTA: id,
                        TITULAR: nombre,
                        CUIL: cuit,
                        "RAZON SOCIAL": razSocial,
                        "COND. IVA": "Error",
                        "ERROR": errorStr,
                        "DIRECCION": "",
                        "ACTIVIDAD": ""
                    }
                    rowsArray.push(row)
                }
            }
        } else {
            const errorStr = verificar.message
            row = {
                CUENTA: id,
                TITULAR: nombre,
                CUIL: cuit,
                "RAZON SOCIAL": "",
                "COND. IVA": "Error",
                "ERROR": errorStr,
                "DIRECCION": "",
                "ACTIVIDAD": "",
            }
            rowsArray.push(row)
        }
    }

    const workbook = XLSX.utils.book_new();
    const Hoja1 = XLSX.utils.json_to_sheet(rowsArray);
    XLSX.utils.book_append_sheet(workbook, Hoja1, "Hoja1", true);
    XLSX.writeFile(workbook, path.join(__dirname, "Files", "Output", "DataCuits.xlsx"));
}

principal();