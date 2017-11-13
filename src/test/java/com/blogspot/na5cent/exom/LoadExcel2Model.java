/*
 * code https://github.com/jittagornp/excel-object-mapping
 */
package com.blogspot.na5cent.exom;

import java.io.File;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import bsh.This;

/**
 * @author redcrow
 */
public class LoadExcel2Model {

    private static final Logger LOG = LoggerFactory.getLogger(LoadExcel2Model.class);

    public static void main(String args[]) throws Throwable {
        File excelFile = new File(LoadExcel2Model.class.getResource("/excel.xlsx").getPath());
        List<Model> items = ExOM.mapFromExcel(excelFile)
                .to(Model.class)
                .map();

        File excelFile1 = new File("/home/vishwanath/Desktop/excel.xlsx");

        ExOM.mapFromExcel(excelFile1).to(Model.class).write(items);
        for (Model item : items) {
            LOG.debug("first name --> {}", item.getFistName());
            LOG.debug("last name --> {}", item.getLastName());
            LOG.debug("age --> {}", item.getAge());
            LOG.debug("birth date --> {}", item.getBirthdate());
            LOG.debug("");
        }
    }
}
