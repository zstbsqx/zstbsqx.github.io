import React, { useEffect, useState } from "react";
import Button from "@material-ui/core/Button";
import Container from "@material-ui/core/Container";
import FilledInput from "@material-ui/core/FilledInput";
import FormControl from "@material-ui/core/FormControl";
import Grid from "@material-ui/core/Grid";
import InputLabel from "@material-ui/core/InputLabel";
import MenuItem from "@material-ui/core/MenuItem";
import Select from "@material-ui/core/Select";
import Stack from "@material-ui/core/Stack";
import { makeStyles } from "@material-ui/styles";
import XLSX from "xlsx";
import nzhcn from "nzh/cn";

const useStyles = makeStyles({
    button: {
        height: '100%',
        contained: true,
    }
});

function ExcelUploader({setWorkbook, setFilename}) {
    const classes = useStyles();
    const [text, setText] = useState("选择要转换的Excel文件");

    const uploadExcel = (event) => {
        if (event.target.files.length === 0) {
            return;
        }
        const file = event.target.files[0];
        setFilename(file.name);
        const reader = new FileReader();
        reader.onload = (event) => {
            const array_buffer = event.target.result;
            const workbook = XLSX.read(array_buffer, {type: 'array'});
            setText(`${file.name} 上传完毕`);
            setWorkbook(workbook);
        };
        setText(`${file.name} 上传中...`);
        reader.readAsArrayBuffer(file);
    }

    return (
        <Grid container spacing={2}>
            <Grid item xs={10}>
                <FilledInput value={text} inputProps={{ style: { padding: '16.5px 14px' } }} readOnly disableUnderline fullWidth />
            </Grid>
            <Grid item xs={2}>
                <Button variant="contained" component="label" className={classes.button} fullWidth>
                    上传
                    <input type="file" accept=".xls,.xlsx" onChange={uploadExcel} hidden />
                </Button>
            </Grid>
        </Grid>
    )
}

function SheetSelector({sheets, onChange}) {
    const [selectedSheetName, setSelectedSheetName] = useState();
    const finalSelectedSheet= sheets.includes(selectedSheetName) ? selectedSheetName : sheets[0];

    useEffect(() => {
        if (selectedSheetName !== finalSelectedSheet) {
            setSelectedSheetName(finalSelectedSheet);
        }
    }, [selectedSheetName, finalSelectedSheet]);

    useEffect(() => {
        if (finalSelectedSheet !== undefined) {
            onChange(finalSelectedSheet);
        }
    }, [finalSelectedSheet, onChange]);

    if (finalSelectedSheet === undefined) {
        return (
            <FormControl fullWidth>
                <InputLabel id="sheet-select-label">工作表</InputLabel>
                <Select labelId="sheet-select-label" label="工作表" value="dummy" disabled>
                    <MenuItem value="dummy" key={0}>没有有效的工作表，请先上传合法的Excel文件</MenuItem>
                </Select>
            </FormControl>
        )
    }

    return (
        <FormControl fullWidth>
            <InputLabel id="sheet-select-label">工作表</InputLabel>
            <Select labelId="sheet-select-label" label="工作表" value={finalSelectedSheet} onChange={(event) => setSelectedSheetName(event.target.value)}>
                {sheets.map((sheet, index) => <MenuItem value={sheet} key={index}>{sheet}</MenuItem>)}
            </Select>
        </FormControl>
    );
}

function ColumnSelector({columns, onChange}) {
    const [selectedColumns, setSelectedColumns] = useState([]);

    useEffect(() => {
        onChange(selectedColumns);
    }, [onChange, selectedColumns]);

    if (columns.length === 0) {
        return (
            <FormControl>
                <InputLabel id="column-select-label">要转换的列</InputLabel>
                <Select labelId="column-select-label" label="要转换的列" multiple value={["dummy"]} disabled>
                    <MenuItem value="dummy" key={0}>没有可用的列，请先选择可用的工作表</MenuItem>
                </Select>
            </FormControl>
        )
    }

    return (
        <FormControl>
            <InputLabel id="column-select-label">要转换的列</InputLabel>
            <Select labelId="column-select-label" label="要转换的列" multiple value={selectedColumns} onChange={(event) => setSelectedColumns(event.target.value)}>
                {columns.map((column) => <MenuItem value={column} key={column.id}>{column.text}</MenuItem>)}
            </Select>
        </FormControl>
    );
}

class TableHeaderGetter {
    constructor() {
        this.columns_map = new Map();
    }

    get(sheet) {
        console.log('start getting table headers');
        const cached_columns = this.columns_map.get(sheet);
        if (cached_columns !== undefined) {
            console.log('Cached!', cached_columns);
            return cached_columns;
        }
        
        console.log('Missed cache!', sheet);
        const columns = [];
        const range = XLSX.utils.decode_range(sheet['!ref']);
        for (let c = range.s.c; c <= range.e.c; ++c) {
            const cell_ref = XLSX.utils.encode_cell({c, r: range.s.r});
            const cell = sheet[cell_ref];
            if (cell !== undefined && cell.w !== undefined) {
                columns.push({id: c, text: cell.w});
            }
        }
        console.log('Result', columns);
        this.columns_map.set(sheet, columns);
        return columns;
    }
}
const table_header_getter = new TableHeaderGetter();

function convertChineseNumToArabic(str) {
    const re = /(^.*?)([零一二三四五六七八九十百千万]+)店$/u;
    const replace = (_, p1, p2) => {
        const arabic_number = nzhcn.decodeS(p2);
        return `${p1}${arabic_number.toString().padStart(2, '0')}店`;
    }
    return str.replace(re, replace);
}

function convertSheet(sheet, columns) {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    for (let column of columns) {
        console.log(`Converting column ${JSON.stringify(column)}`);
        for (let r = range.s.r + 1; r <= range.e.r; ++r) {
            const cell_ref = XLSX.utils.encode_cell({c: column.id, r});
            const cell = sheet[cell_ref];
            if (cell !== undefined && cell.w !== undefined) {
                cell.v = convertChineseNumToArabic(cell.w)
                console.log(`${cell.w} -> ${cell.v}`);
                delete cell.w;
            }
        }
    }
}

function App() {
    const [filename, setFilename] = useState();
    const [workbook, setWorkbook] = useState();
    const [selectedSheetName, setSelectedSheetName] = useState();
    const [selectedColumns, setSelectedColumns] = useState([]);

    const columns = selectedSheetName === undefined ? [] : table_header_getter.get(workbook.Sheets[selectedSheetName]);

    const convertAndDownload = () => {
        console.log(selectedSheetName);
        const sheet = workbook.Sheets[selectedSheetName];
        convertSheet(sheet, selectedColumns);
        const binary = XLSX.write(workbook, { type: 'array' });
        const blob = new Blob([binary], { type: 'application/vnd.ms-excel'});
        const url = URL.createObjectURL(blob);

        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download= filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    return (
        <Container>
            <Stack spacing={2}>
                <ExcelUploader setWorkbook={setWorkbook} setFilename={setFilename}/>
                <SheetSelector sheets={workbook === undefined ? [] : workbook.SheetNames} onChange={setSelectedSheetName} />
                <ColumnSelector columns={columns} onChange={setSelectedColumns} />
                <Button variant="contained" onClick={convertAndDownload} disabled={[workbook, selectedSheetName, selectedColumns].some((x) => x === undefined)}>转换并下载</Button>
            </Stack>
        </Container>
    );
}

export default App;
