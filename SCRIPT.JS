document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('fileInput');
    const showListBtn = document.getElementById('showListBtn');
    const formGroupsBtn = document.getElementById('formGroupsBtn');
    const studentListUl = document.getElementById('studentListUl');
    const groupsContainer = document.getElementById('groupsContainer');
    const numGruposInput = document.getElementById('numGrupos');
    const numMujeresInput = document.getElementById('numMujeres');
    const numHombresInput = document.getElementById('numHombres');
    const fileInfo = document.getElementById('fileInfo');
    const errorMessage = document.getElementById('error-message');

    let alumnos = [];
    let archivoCargado = false;

    fileInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];

        if (!file) {
            fileInfo.textContent = 'Ningún archivo seleccionado.';
            showListBtn.disabled = true;
            formGroupsBtn.disabled = true;
            archivoCargado = false;
            return;
        }

        fileInfo.textContent = `Archivo seleccionado: ${file.name}`;

        try {
            alumnos = await procesarArchivo(file);
            showListBtn.disabled = false;
            formGroupsBtn.disabled = false;
            archivoCargado = true;
            errorMessage.textContent = ''; // Limpiar mensaje de error
        } catch (error) {
            console.error("Error al procesar el archivo:", error);
            errorMessage.textContent = 'Error al procesar el archivo. Verifique el formato.';
            showListBtn.disabled = true;
            formGroupsBtn.disabled = true;
            archivoCargado = false;
            alumnos = []; // Limpiar los datos
        }
    });

    showListBtn.addEventListener('click', () => {
        if (!archivoCargado) {
            alert('Por favor, cargue un archivo primero.');
            return;
        }
        mostrarListaAlumnos(alumnos);
    });

    formGroupsBtn.addEventListener('click', () => {
        if (!archivoCargado) {
            alert('Por favor, cargue un archivo primero.');
            return;
        }
        const numGrupos = parseInt(numGruposInput.value, 10);
        const numMujeres = parseInt(numMujeresInput.value, 10);
        const numHombres = parseInt(numHombresInput.value, 10);

        if (isNaN(numGrupos) || numGrupos <= 0) {
            alert('Por favor, ingrese un número válido de grupos.');
            return;
        }

        if (isNaN(numMujeres) || isNaN(numHombres) || numMujeres < 0 || numHombres < 0) {
            alert('Por favor, ingrese un número válido de mujeres y hombres por grupo.');
            return;
        }

        formarGrupos(alumnos, numGrupos, numMujeres, numHombres);
    });

    function procesarArchivo(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();

            reader.onload = (e) => {
                const data = e.target.result;
                const workbook = XLSX.read(data, { type: 'binary' });
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                // Convertir a formato CSV (simplemente uniendo las filas con comas)
                const csvData = jsonData.map(row => row.join(';')).join('\n');

                // Parsea el CSV
                const alumnos = parseCSV(csvData);

                if (alumnos && alumnos.length > 0) {
                    resolve(alumnos);
                } else {
                    reject('No se pudieron extraer los datos de alumnos del archivo.');
                }
            };

            reader.onerror = (error) => {
                reject(error);
            };

            reader.readAsBinaryString(file);
        });
    }


    function parseCSV(csvText) {
        const lineas = csvText.split('\n');
        const alumnos = [];

        // Asume que la primera línea son los encabezados
        let encabezados = lineas[0].split(';');

        // Eliminar el BOM si está presente
        if (encabezados[0].charCodeAt(0) === 65279) {
            encabezados[0] = encabezados[0].substring(1);
        }

        // Encuentra los índices de las columnas necesarias (ignorando mayúsculas/minúsculas y espacios)
        const idxApellidosNombres = findColumnIndex(encabezados, 'ApellidosNombres');
        const idxGenero = findColumnIndex(encabezados, 'Genero');

        if (idxApellidosNombres === -1 || idxGenero === -1) {
            console.error("No se encontraron las columnas 'ApellidosNombres' o 'Genero'.");
            return [];
        }

        for (let i = 1; i < lineas.length; i++) {  // i = 1 para saltar la línea de encabezado
            const valores = lineas[i].split(';');
            // Verifica que la línea tenga suficientes columnas antes de acceder a ellas
            if (valores.length > Math.max(idxApellidosNombres, idxGenero)) {
                const apellidosNombres = valores[idxApellidosNombres].trim();
                const genero = valores[idxGenero].trim().toUpperCase(); // Convertir a mayúsculas para comparación

                // Valida que los datos no estén vacíos
                if (apellidosNombres && genero) {
                    alumnos.push({
                        apellidosNombres: apellidosNombres,
                        genero: genero
                    });
                }
            }
        }
        return alumnos;
    }

    function findColumnIndex(headerRow, columnName) {
        const normalizedColumnName = columnName.trim().toLowerCase();
        for (let i = 0; i < headerRow.length; i++) {
            if (headerRow[i].trim().toLowerCase() === normalizedColumnName) {
                return i;
            }
        }
        return -1; // Retorna -1 si no se encuentra la columna
    }


    function mostrarListaAlumnos(alumnos) {
        studentListUl.innerHTML = '';
        const alumnosOrdenados = [...alumnos].sort((a, b) =>
            a.apellidosNombres.localeCompare(b.apellidosNombres)
        );

        alumnosOrdenados.forEach(alumno => {
            const li = document.createElement('li');
            li.textContent = `${alumno.apellidosNombres} (${alumno.genero})`;
            studentListUl.appendChild(li);
        });
    }


    function formarGrupos(alumnos, numGrupos, numMujeresPorGrupo, numHombresPorGrupo) {
        // Validación adicional para asegurar que haya suficientes alumnos
        const totalAlumnosNecesarios = numGrupos * (numMujeresPorGrupo + numHombresPorGrupo);
        if (alumnos.length < totalAlumnosNecesarios) {
            alert('No hay suficientes alumnos para formar los grupos con la configuración especificada.');
            return;
        }

        const mujeres = alumnos.filter(alumno => alumno.genero === 'F');
        const hombres = alumnos.filter(alumno => alumno.genero === 'M');

        if (mujeres.length < numGrupos * numMujeresPorGrupo || hombres.length < numGrupos * numHombresPorGrupo) {
            alert('No hay suficientes mujeres u hombres para formar los grupos con la configuración especificada.');
            return;
        }

        const grupos = [];
        const alumnosAsignados = new Set(); // Para evitar nombres repetidos

        for (let i = 0; i < numGrupos; i++) {
            const grupo = {
                id: i + 1,
                integrantes: []
            };

            // Agregar mujeres al grupo
            for (let j = 0; j < numMujeresPorGrupo; j++) {
                let mujerAsignada = null;
                let intentos = 0;

                // Intenta encontrar una mujer no asignada, con un límite de intentos
                while (intentos < mujeres.length) {
                    const indiceMujer = Math.floor(Math.random() * mujeres.length);
                    const posibleMujer = mujeres[indiceMujer];

                    if (!alumnosAsignados.has(posibleMujer.apellidosNombres)) {
                        mujerAsignada = posibleMujer;
                        break;
                    }
                    intentos++;
                }

                if (mujerAsignada) {
                    grupo.integrantes.push(mujerAsignada);
                    alumnosAsignados.add(mujerAsignada.apellidosNombres);
                } else {
                    console.warn(`No se pudo asignar una mujer al grupo ${i + 1} después de varios intentos.`);
                    // Aquí puedes implementar lógica adicional si no se puede cumplir con la cuota de género.
                }
            }


            // Agregar hombres al grupo
            for (let j = 0; j < numHombresPorGrupo; j++) {
                let hombreAsignado = null;
                let intentos = 0;

                // Intenta encontrar un hombre no asignado
                while (intentos < hombres.length) {
                    const indiceHombre = Math.floor(Math.random() * hombres.length);
                    const posibleHombre = hombres[indiceHombre];

                    if (!alumnosAsignados.has(posibleHombre.apellidosNombres)) {
                        hombreAsignado = posibleHombre;
                        break;
                    }
                    intentos++;
                }

                if (hombreAsignado) {
                    grupo.integrantes.push(hombreAsignado);
                    alumnosAsignados.add(hombreAsignado.apellidosNombres);
                } else {
                    console.warn(`No se pudo asignar un hombre al grupo ${i + 1} después de varios intentos.`);
                    // Aquí puedes implementar lógica adicional si no se puede cumplir con la cuota de género.
                }
            }


            grupos.push(grupo);
        }

        mostrarGrupos(grupos);
    }



    function mostrarGrupos(grupos) {
        groupsContainer.innerHTML = '';

        grupos.forEach(grupo => {
            const groupDiv = document.createElement('div');
            groupDiv.classList.add('group');

            const h3 = document.createElement('h3');
            h3.textContent = `Grupo ${grupo.id}`;
            groupDiv.appendChild(h3);

            const ul = document.createElement('ul');
            grupo.integrantes.forEach(integrante => {
                const li = document.createElement('li');
                li.textContent = `${integrante.apellidosNombres} (${integrante.genero})`;
                ul.appendChild(li);
            });
            groupDiv.appendChild(ul);

            groupsContainer.appendChild(groupDiv);
        });
    }
});