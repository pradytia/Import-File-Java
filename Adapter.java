public static response ImportExcelDistrict(Workbook workbook, modelParam param) {

        String user = param.created_by;
        String sql = "";
        String function_name = Thread.currentThread().getStackTrace()[1].getMethodName();
        boolean resultExec = false;
        List<modelParam> _modelParamList = new ArrayList<modelParam>();
        response _response = new respons();
        ArrayList<String> idList = new ArrayList<>();

        try {
            
//            int getSheet = workbook.getNumberOfSheets();
//            System.out.println(getSheet);
            Sheet firstSheet = workbook.getSheetAt(0);
            Iterator<Row> _rowIterator = firstSheet.iterator();
            ArrayList<String> errorList = new ArrayList<String>();
            ArrayList<Integer> errorRow = new ArrayList<Integer>();
            int errorIndicator = 0;
            int columnIndicator = 1;
            mdlErrorExcel _mdlErrorExcel = new mdlErrorExcel();


            //Skipping Excel header => looping into(row)
            while (_rowIterator.hasNext()) {
                Row currentRow = _rowIterator.next();
                int rowIndex = currentRow.getRowNum();

                if (rowIndex == 0) {
                    continue;
                }

                Iterator<Cell> _cellIterator = currentRow.cellIterator();
                modelParam _modelParam = new modelParam();

                //looping ke samping(column)
                while(_cellIterator.hasNext()) {
                    Cell _nextCell = _cellIterator.next();
                    int columnIndex = _nextCell.getColumnIndex();

                    //process get value from excel and set into model

                    if(columnIndex == 0){

                        if(_nextCell.getCellType() == CellType.STRING){
                            String valueID = _nextCell.getStringCellValue();
                            _modelParam.setID(valueID);
                        }else if(_nextCell.getCellType() == CellType.NUMERIC){
                            String valueID = NumberToTextConverter.toText(_nextCell.getNumericCellValue());
                            _modelParam.setID(valueID);
                        }

                    }else if(columnIndex == 1){
                        if(_nextCell.getCellType() == CellType.STRING){
                            String valueRegionID = _nextCell.getStringCellValue();
                            _modelParam.setFK(valueFKID);
                        }else if(_nextCell.getCellType() == CellType.NUMERIC){
                            String valueFKID = NumberToTextConverter.toText(_nextCell.getNumericCellValue());
                            _modelParam.setFK(valueFKID);
                        }else {

                            errorIndicator ++;
                        }
                    }else if(columnIndex == 2){
                        if(_nextCell.getCellType() == CellType.STRING){
                            String valueName = _nextCell.getStringCellValue();
                            _modelParam.setName(valueName);
                        }else if(_nextCell.getCellType() == CellType.NUMERIC){
                            String valueName = NumberToTextConverter.toText(_nextCell.getNumericCellValue());
                            _modelParam.setName(valueName);
                        }
                    }
                }
                _modelParamList.add(_modelParam);
                idList.add(_modelParam.ID);
            }
            workbook.close();

            for(int i = 0; i < idList.size(); i++) {

                if (idList.get(i) == null) {
                    errorRow.add(i + 1);
                    errorIndicator ++;
                }

            }

            if(errorIndicator > 0){

                _respons.Response = errorRow;
                _respons.Status = false;

            } else {

                //INSERT INTO DATABASE

                for (modelParam _modelParam : _modelParamList) {

                    List<mdlQueryExecute> _mdlQueryExecuteList = new ArrayList<mdlQueryExecute>();

                    sql = "{}"; //  ==> YOU CAN INPUT YOUR QUERY HERE

                    _mdlQueryExecuteList.add(QueryAdapter.QueryParam("string", _modelParam.ID));
                    _mdlQueryExecuteList.add(QueryAdapter.QueryParam("string", _modelParam.FK));
                    _mdlQueryExecuteList.add(QueryAdapter.QueryParam("string", _modelParam.Name));
                    _mdlQueryExecuteList.add(QueryAdapter.QueryParam("string", param.created_by));

                    _respons = QueryAdapter.QueryManipulateWithDB2(sql, _mdlQueryExecuteList, function_name, user, Globals.dbName);
                }
            }


        }catch (Exception e){
            core.LogAdapter.InsertLogExc(e.toString(), "UploadExcel", sql, user);
            e.getMessage();
        }
        return _respons;

    }