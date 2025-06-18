
CREATE TABLE `bill` (
  `id` int NOT NULL AUTO_INCREMENT,
  `userid` int DEFAULT NULL,
  `doctorid` int DEFAULT NULL,
  `paymoney` decimal(10,2) DEFAULT NULL,
  `paytime` datetime DEFAULT NULL,
  `state` int DEFAULT NULL COMMENT '订单状态：1已支付 2未支付 ',
  `billcode` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL,
  PRIMARY KEY (`id`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=82 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci ROW_FORMAT=DYNAMIC;

CREATE TABLE `chat_message` (
  `id` int NOT NULL AUTO_INCREMENT COMMENT '每条消息的唯一标识符',
  `senderid` int DEFAULT NULL COMMENT '发送者的id',
  `receiverid` int DEFAULT NULL COMMENT '接收者的id',
  `messageContent` json DEFAULT NULL COMMENT '聊天的消息内容，用json格式',
  `sendtime` timestamp(3) NULL DEFAULT NULL COMMENT '发送的时间',
  `status` int DEFAULT NULL COMMENT '消息的状态，如已读，未读',
  `isdeleted` int DEFAULT '0' COMMENT '标记消息是否被删除，0代表未删除，1代表已删除（默认为0）',
  PRIMARY KEY (`id`) USING BTREE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci ROW_FORMAT=DYNAMIC;

CREATE TABLE `doctor` (
  `doctorid` int NOT NULL AUTO_INCREMENT COMMENT '医生认证id主键',
  `userid` int DEFAULT NULL COMMENT '关联的用户表外键id',
  `licensenumber` varchar(50) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL COMMENT '医师执业证号',
  `licensepicture` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL COMMENT '执业证照片',
  `infopicture` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL COMMENT '个人照片',
  `introduce` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL COMMENT '个人介绍',
  `applydata` datetime DEFAULT NULL COMMENT '审核通过时间',
  `rejectreason` varchar(100) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL COMMENT '拒绝原因',
  `applystatus` int DEFAULT '2' COMMENT '申请状态（0失败，1成功，2申请中）',
  `price` decimal(10,2) DEFAULT NULL COMMENT '咨询价格',
  PRIMARY KEY (`doctorid`) USING BTREE,
  UNIQUE KEY `userid_unique` (`userid`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=54 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci ROW_FORMAT=DYNAMIC;

CREATE TABLE `doctor_income` (
  `id` int NOT NULL AUTO_INCREMENT,
  `doctorid` int DEFAULT NULL,
  `userid` int DEFAULT NULL,
  `money` decimal(10,2) DEFAULT NULL,
  `time` datetime DEFAULT NULL COMMENT '消费时间',
  `point` decimal(10,2) DEFAULT NULL COMMENT '分成',
  `ispay` int DEFAULT NULL COMMENT '是否付款',
  PRIMARY KEY (`id`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=82 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci ROW_FORMAT=DYNAMIC;

CREATE TABLE `doctor_major` (
  `id` int NOT NULL AUTO_INCREMENT,
  `doctorid` int DEFAULT NULL,
  `majorid` int DEFAULT NULL,
  PRIMARY KEY (`id`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=92 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci ROW_FORMAT=DYNAMIC;


CREATE TABLE `major` (
  `majorid` int NOT NULL AUTO_INCREMENT,
  `majorname` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL,
  PRIMARY KEY (`majorid`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=12 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci ROW_FORMAT=DYNAMIC;

CREATE TABLE `manager` (
  `managerid` int NOT NULL AUTO_INCREMENT,
  `managername` varchar(10) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL,
  `username` varchar(32) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL,
  `password` varchar(32) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL,
  `roleid` int DEFAULT NULL,
  PRIMARY KEY (`managerid`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=39 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci ROW_FORMAT=DYNAMIC;

CREATE TABLE `module` (
  `moduleid` int NOT NULL,
  `moduleurl` varchar(50) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL COMMENT '模块路径',
  PRIMARY KEY (`moduleid`) USING BTREE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci ROW_FORMAT=DYNAMIC;

CREATE TABLE `q_a` (
  `id` int NOT NULL AUTO_INCREMENT,
  `question` varchar(125) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL COMMENT '问题',
  `answer` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL COMMENT '答案',
  `doctorid` int DEFAULT NULL,
  `qadate` datetime DEFAULT NULL,
  PRIMARY KEY (`id`) USING BTREE
) ENGINE=InnoDB AUTO_INCREMENT=26 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci ROW_FORMAT=DYNAMIC;

CREATE TABLE `role` (
  `roleid` int NOT NULL,
  `rolename` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL,
  PRIMARY KEY (`roleid`) USING BTREE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci ROW_FORMAT=DYNAMIC;

CREATE TABLE `role_module` (
  `id` int NOT NULL,
  `roleid` int DEFAULT NULL,
  `moduleid` int DEFAULT NULL,
  PRIMARY KEY (`id`) USING BTREE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci ROW_FORMAT=DYNAMIC;

CREATE TABLE `user` (
  `userid` int NOT NULL AUTO_INCREMENT COMMENT '用户表id主键',
  `username` varchar(32) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci NOT NULL COMMENT '用户名',
  `password` varchar(32) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL COMMENT '密码',
  `phone` varchar(13) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL COMMENT '手机号',
  `headphoto` varchar(255) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL COMMENT '头像路径',
  `idcard` varchar(30) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL COMMENT '身份证号',
  `gender` char(2) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL COMMENT '性别',
  `borndate` date DEFAULT NULL COMMENT '出生日期',
  `createdtime` datetime DEFAULT NULL COMMENT '注册时间',
  `roleid` int unsigned DEFAULT '1001' COMMENT '1001普通用户2002健康师',
  `name` varchar(20) CHARACTER SET utf8mb4 COLLATE utf8mb4_general_ci DEFAULT NULL COMMENT '昵称',
  `status` int DEFAULT NULL COMMENT '1在线0不在线',
  PRIMARY KEY (`userid`) USING BTREE,
  UNIQUE KEY `username` (`username`)
) ENGINE=InnoDB AUTO_INCREMENT=1047 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci ROW_FORMAT=DYNAMIC;

CREATE TABLE `user_role` (
  `id` int NOT NULL,
  `userid` int DEFAULT NULL,
  `roleid` int DEFAULT NULL,
  PRIMARY KEY (`id`) USING BTREE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci ROW_FORMAT=DYNAMIC;
