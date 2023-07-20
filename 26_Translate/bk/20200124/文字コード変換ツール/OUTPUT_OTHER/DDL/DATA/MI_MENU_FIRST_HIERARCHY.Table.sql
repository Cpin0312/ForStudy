﻿DELETE FROM [MI_MENU_FIRST_HIERARCHY]
GO
INSERT [MI_MENU_FIRST_HIERARCHY] ([FIRST_HIERARCHY_ID], [FIRST_HIERARCHY_NAME], [RIGHT_OPERATE], [SHOW_ORDER], [REG_ID], [REG_DATE], [UPD_ID], [UPD_DATE]) VALUES (N'FD', N'文檔功能表', N'OPC131,OPC141,OPC101,OPC102,OPC031,OPC041,OPC001,OPC002,TMF001,TMF004,HDK001,HSE001', 2, N'pi', CAST(N'2015-08-01T00:00:00.000' AS DateTime), NULL, NULL)
INSERT [MI_MENU_FIRST_HIERARCHY] ([FIRST_HIERARCHY_ID], [FIRST_HIERARCHY_NAME], [RIGHT_OPERATE], [SHOW_ORDER], [REG_ID], [REG_DATE], [UPD_ID], [UPD_DATE]) VALUES (N'MN', N'管理功能表', N'HSE001', 3, N'pi', CAST(N'2015-08-01T00:00:00.000' AS DateTime), NULL, NULL)
INSERT [MI_MENU_FIRST_HIERARCHY] ([FIRST_HIERARCHY_ID], [FIRST_HIERARCHY_NAME], [RIGHT_OPERATE], [SHOW_ORDER], [REG_ID], [REG_DATE], [UPD_ID], [UPD_DATE]) VALUES (N'US', N'作業清單表', N'OPC131,OPC141,OPC101,OPC102,OPC103,INF101,INF102,OPC121,OPC122,OPC111,OPC112,OPC031,OPC041,OPC001,OPC002,OPC003,INF001,INF002,OPC011,OPC012,TMF001,TMF002,TMF003,TMF004,CCT001,HDK001,HSE001', 1, N'pi', CAST(N'2015-08-01T00:00:00.000' AS DateTime), NULL, NULL)
