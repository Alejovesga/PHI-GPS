/*
 * Copyright 2016 - 2022 Anton Tananaev (anton@traccar.org)
 * Copyright 2016 Andrey Kunitsyn (andrey@traccar.org)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.traccar.reports;

import org.apache.poi.ss.util.WorkbookUtil;
import org.traccar.config.Config;
import org.traccar.config.Keys;
import org.traccar.helper.model.DeviceUtil;
import org.traccar.helper.model.PositionUtil;
import org.traccar.model.Device;
import org.traccar.model.Group;
import org.traccar.model.Position;
import org.traccar.reports.common.ReportUtils;
import org.traccar.reports.model.DeviceReportSection;
import org.traccar.storage.Storage;
import org.traccar.storage.StorageException;
import org.traccar.storage.query.Columns;
import org.traccar.storage.query.Condition;
import org.traccar.storage.query.Request;

import javax.inject.Inject;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;

public class DevicesReportProvider {

    private final Config config;
    private final ReportUtils reportUtils;
    private final Storage storage;


    @Inject
    public DevicesReportProvider(Config config, ReportUtils reportUtils, Storage storage) {
        this.config = config;
        this.reportUtils = reportUtils;
        this.storage = storage;
    }

    public Collection<Position> getObjects(long userId, Collection<Long> deviceIds, Collection<Long> groupIds,
            Date from, Date to) throws StorageException {
        reportUtils.checkPeriodLimit(from, to);

        ArrayList<Position> result = new ArrayList<>();
        for (Device device: DeviceUtil.getAccessibleDevices(storage, userId, deviceIds, groupIds)) {
            result.addAll(PositionUtil.getPositions(storage, device.getId(), from, to));
        }
        return result;
    }

    public void getDeviceReportExcel(OutputStream outputStream,
            long userId) throws StorageException, IOException {

        ArrayList<DeviceReportSection> devicesReport = new ArrayList<>();
        ArrayList<String> sheetNames = new ArrayList<>();
        String sheetName = "Devices Report";  // Nombre personalizado para la hoja de cálculo
        sheetNames.add(WorkbookUtil.createSafeSheetName(sheetName));
        for (Device device: DeviceUtil.getDevices(storage, userId)) {
            var devices = DeviceUtil.getDevices(storage, userId);
            DeviceReportSection deviceReport = new DeviceReportSection();
            deviceReport.setDeviceName(device.getName());
            if (device.getGroupId() > 0) {
                Group group = storage.getObject(Group.class, new Request(
                        new Columns.All(), new Condition.Equals("id", device.getGroupId())));
                if (group != null) {
                    deviceReport.setGroupName(group.getName());
                }
            }
            deviceReport.setObjects(devices);
            devicesReport.add(deviceReport);
        }

        File file = Paths.get(config.getString(Keys.TEMPLATES_ROOT), "export", "devices.xlsx").toFile();
        try (InputStream inputStream = new FileInputStream(file)) {
            var context = reportUtils.initializeContext(userId);
            context.putVar("devices", devicesReport);
            context.putVar("sheetNames", sheetNames);
            reportUtils.processTemplateWithSheets(inputStream, outputStream, context);
        }
    }
}
