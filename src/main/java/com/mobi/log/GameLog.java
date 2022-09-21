package com.mobi.log;


import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class GameLog {
    private static final Logger logger = LoggerFactory.getLogger(GameLog.class);
    public static void LogInfo(String strLog, Object... args) {
        logger.info(strLog, args);
    }

    public static void LogError(String strLog, Object... args) {
        logger.error(strLog, args);
    }

    public static void LogException(String strLog, Object... args) {
        logger.error(strLog, args);
    }
}
