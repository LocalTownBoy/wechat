CREATE TABLE `papers` (
    `id` int NOT NULL AUTO_INCREMENT,
    `title` varchar(255) NOT NULL COMMENT '论文标题',
    `author` varchar(255) NOT NULL COMMENT '作者',
    `section` varchar(255) NOT NULL COMMENT '章节',
    `created_at` datetime(6) NOT NULL COMMENT '创建时间',
    `url` varchar(512) DEFAULT NULL COMMENT 'URL',
    PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;