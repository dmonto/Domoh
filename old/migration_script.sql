-- ----------------------------------------------------------------------------
-- MySQL Workbench Migration
-- Migrated Schemata: coachai
-- Source Schemata: coachai
-- Created: Wed Sep 18 18:30:04 2019
-- Workbench Version: 8.0.15
-- ----------------------------------------------------------------------------

SET FOREIGN_KEY_CHECKS = 0;

-- ----------------------------------------------------------------------------
-- Schema coachai
-- ----------------------------------------------------------------------------
DROP SCHEMA IF EXISTS `coachai` ;
CREATE SCHEMA IF NOT EXISTS `coachai` ;

-- ----------------------------------------------------------------------------
-- Table coachai.eahead
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `coachai`.`eahead` (
  `EASet` VARCHAR(40) NULL DEFAULT NULL,
  `Fichero` VARCHAR(200) NULL DEFAULT NULL,
  `Symbol` VARCHAR(40) NULL DEFAULT NULL,
  `Period` VARCHAR(40) NULL DEFAULT NULL,
  `Comienzo` VARCHAR(40) NULL DEFAULT NULL,
  `Final` VARCHAR(40) NULL DEFAULT NULL,
  `Bars` VARCHAR(40) NULL DEFAULT NULL,
  `Ticks` VARCHAR(40) NULL DEFAULT NULL,
  `ModQual` VARCHAR(40) NULL DEFAULT NULL,
  `Mismatch` VARCHAR(40) NULL DEFAULT NULL,
  `IniDepo` VARCHAR(40) NULL DEFAULT NULL,
  `Spread` VARCHAR(40) NULL DEFAULT NULL,
  `TotNet` VARCHAR(40) NULL DEFAULT NULL,
  `GrossProf` VARCHAR(40) NULL DEFAULT NULL,
  `GrossLoss` VARCHAR(40) NULL DEFAULT NULL,
  `ProfFactor` VARCHAR(40) NULL DEFAULT NULL,
  `ExpPayoff` VARCHAR(40) NULL DEFAULT NULL,
  `AbsDraw` VARCHAR(40) NULL DEFAULT NULL,
  `MaxDraw` VARCHAR(40) NULL DEFAULT NULL,
  `MaxDrawPerc` VARCHAR(40) NULL DEFAULT NULL,
  `RelDraw` VARCHAR(40) NULL DEFAULT NULL,
  `RelDrawPerc` VARCHAR(40) NULL DEFAULT NULL,
  `TotTrades` VARCHAR(40) NULL DEFAULT NULL,
  `Shorts` VARCHAR(40) NULL DEFAULT NULL,
  `ShortPerc` VARCHAR(40) NULL DEFAULT NULL,
  `Longs` VARCHAR(40) NULL DEFAULT NULL,
  `LongPerc` VARCHAR(40) NULL DEFAULT NULL,
  `Profitable` VARCHAR(40) NULL DEFAULT NULL,
  `ProfitablePerc` VARCHAR(40) NULL DEFAULT NULL,
  `Losing` VARCHAR(40) NULL DEFAULT NULL,
  `LosingPerc` VARCHAR(40) NULL DEFAULT NULL,
  `LargestProfit` VARCHAR(40) NULL DEFAULT NULL,
  `LargestLoss` VARCHAR(40) NULL DEFAULT NULL,
  `AvgProfit` VARCHAR(40) NULL DEFAULT NULL,
  `AvgLoss` VARCHAR(40) NULL DEFAULT NULL,
  `WinDealStreak` VARCHAR(40) NULL DEFAULT NULL,
  `WinDealStreakCash` VARCHAR(40) NULL DEFAULT NULL,
  `LosDealStreak` VARCHAR(40) NULL DEFAULT NULL,
  `LosDealStreakCash` VARCHAR(40) NULL DEFAULT NULL,
  `WinCashStreak` VARCHAR(40) NULL DEFAULT NULL,
  `WinCashStreakDeals` VARCHAR(40) NULL DEFAULT NULL,
  `LosCashStreak` VARCHAR(40) NULL DEFAULT NULL,
  `LosCashStreakDeals` VARCHAR(40) NULL DEFAULT NULL,
  `AvgWinStreak` VARCHAR(40) NULL DEFAULT NULL,
  `AvgLosStreak` VARCHAR(40) NULL DEFAULT NULL)
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8mb4
ROW_FORMAT = COMPRESSED;

-- ----------------------------------------------------------------------------
-- Table coachai.eaparams
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `coachai`.`eaparams` (
  `EASet` VARCHAR(40) NULL DEFAULT NULL,
  `Param` VARCHAR(40) NULL DEFAULT NULL,
  `Value` VARCHAR(80) NULL DEFAULT NULL)
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8mb4;

-- ----------------------------------------------------------------------------
-- Table coachai.easet
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `coachai`.`easet` (
  `Indice` INT(11) NOT NULL,
  `EASetName` VARCHAR(100) NULL DEFAULT NULL,
  `Time` VARCHAR(20) NULL DEFAULT NULL,
  `Type` VARCHAR(20) NULL DEFAULT NULL,
  `Order` INT(11) NULL DEFAULT NULL,
  `Size` DOUBLE NULL DEFAULT NULL,
  `Price` VARCHAR(20) NULL DEFAULT NULL,
  `S / L` VARCHAR(20) NULL DEFAULT NULL,
  `T / P` VARCHAR(20) NULL DEFAULT NULL,
  `Profit` VARCHAR(20) NULL DEFAULT NULL,
  `Balance` VARCHAR(20) NULL DEFAULT NULL)
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8mb4;

-- ----------------------------------------------------------------------------
-- Table coachai.easetiters
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `coachai`.`easetiters` (
  `SetName` VARCHAR(100) NULL DEFAULT NULL,
  `index` BIGINT(20) NULL DEFAULT NULL,
  `Iteration` BIGINT(20) NULL DEFAULT NULL,
  `Cycle` BIGINT(20) NULL DEFAULT NULL,
  `Status` TEXT NULL DEFAULT NULL,
  `EA` TEXT NULL DEFAULT NULL,
  `Symbol` TEXT NULL DEFAULT NULL,
  `Period` BIGINT(20) NULL DEFAULT NULL,
  `Spread` BIGINT(20) NULL DEFAULT NULL,
  `ALL.FirstDate` TEXT NULL DEFAULT NULL,
  `ALL.LastDate` TEXT NULL DEFAULT NULL,
  `BAS.FirstDate` TEXT NULL DEFAULT NULL,
  `BAS.LastDate` TEXT NULL DEFAULT NULL,
  `ALLOOS.FirstDate` TEXT NULL DEFAULT NULL,
  `ALLOOS.LastDate` TEXT NULL DEFAULT NULL,
  `OOS1.FirstDate` TEXT NULL DEFAULT NULL,
  `OOS1.LastDate` TEXT NULL DEFAULT NULL,
  `OOS2.FirstDate` TEXT NULL DEFAULT NULL,
  `OOS2.LastDate` TEXT NULL DEFAULT NULL,
  `OOS3.FirstDate` TEXT NULL DEFAULT NULL,
  `OOS3.LastDate` TEXT NULL DEFAULT NULL,
  `OOS4.FirstDate` TEXT NULL DEFAULT NULL,
  `OOS4.LastDate` TEXT NULL DEFAULT NULL,
  `OOS5.FirstDate` TEXT NULL DEFAULT NULL,
  `OOS5.LastDate` TEXT NULL DEFAULT NULL,
  `OOS6.FirstDate` TEXT NULL DEFAULT NULL,
  `OOS6.LastDate` TEXT NULL DEFAULT NULL,
  `FirstOp` TEXT NULL DEFAULT NULL,
  `LastOp` TEXT NULL DEFAULT NULL,
  `Param Name` TEXT NULL DEFAULT NULL,
  `Param Value` DOUBLE NULL DEFAULT NULL,
  `FO Name` TEXT NULL DEFAULT NULL,
  INDEX `ix_easetiters_index` (`index` ASC) VISIBLE)
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8mb4;

-- ----------------------------------------------------------------------------
-- Table coachai.easetparams
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `coachai`.`easetparams` (
  `index` INT(11) NULL DEFAULT NULL,
  `SetName` VARCHAR(100) NULL DEFAULT NULL,
  `SetIndex` INT(11) NULL DEFAULT NULL,
  `Param` VARCHAR(100) NULL DEFAULT NULL,
  `Value` VARCHAR(200) NULL DEFAULT NULL,
  INDEX `ix_easetparams_index` (`index` ASC) VISIBLE,
  INDEX `ix_setname_param` (`SetName` ASC, `Param` ASC) VISIBLE)
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8mb4;

-- ----------------------------------------------------------------------------
-- Table coachai.easettrades
-- ----------------------------------------------------------------------------
CREATE TABLE IF NOT EXISTS `coachai`.`easettrades` (
  `?EASetName` VARCHAR(100) NULL DEFAULT NULL,
  `Order` INT(11) NULL DEFAULT NULL,
  `IndiceOpen` INT(11) NULL DEFAULT NULL,
  `IndiceClose` INT(11) NULL DEFAULT NULL,
  `TimeOpen` VARCHAR(20) NULL DEFAULT NULL,
  `TimeClose` VARCHAR(20) NULL DEFAULT NULL,
  `TypeOpen` VARCHAR(20) NULL DEFAULT NULL,
  `TypeClose` VARCHAR(20) NULL DEFAULT NULL,
  `Size` DOUBLE NULL DEFAULT NULL,
  `PriceOpen` DOUBLE NULL DEFAULT NULL,
  `PriceClose` DOUBLE NULL DEFAULT NULL,
  `SLOpen` DOUBLE NULL DEFAULT NULL,
  `SLHigh` DOUBLE NULL DEFAULT NULL,
  `SLLow` DOUBLE NULL DEFAULT NULL,
  `SLClose` DOUBLE NULL DEFAULT NULL,
  `TPOpen` DOUBLE NULL DEFAULT NULL,
  `TPHigh` DOUBLE NULL DEFAULT NULL,
  `TPLow` DOUBLE NULL DEFAULT NULL,
  `TPClose` DOUBLE NULL DEFAULT NULL,
  `Figura03` VARCHAR(1) NULL DEFAULT NULL,
  `Figura04` VARCHAR(1) NULL DEFAULT NULL,
  `Figura05` VARCHAR(1) NULL DEFAULT NULL,
  `Figura06` VARCHAR(1) NULL DEFAULT NULL,
  `Figura07` VARCHAR(1) NULL DEFAULT NULL,
  `Figura08` VARCHAR(1) NULL DEFAULT NULL,
  `Figura09` VARCHAR(1) NULL DEFAULT NULL,
  `Figura10` VARCHAR(1) NULL DEFAULT NULL,
  `Modifications` MEDIUMTEXT NULL DEFAULT NULL,
  `Profit` DOUBLE NULL DEFAULT NULL,
  `Balance` DOUBLE NULL DEFAULT NULL,
  INDEX `ixsettrades` (`?EASetName` ASC, `IndiceOpen` ASC) VISIBLE)
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8mb4;
SET FOREIGN_KEY_CHECKS = 1;
