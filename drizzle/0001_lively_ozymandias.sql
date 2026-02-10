CREATE TABLE `drafts` (
	`id` int AUTO_INCREMENT NOT NULL,
	`reportId` int,
	`userId` int NOT NULL,
	`draftData` json,
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	`updatedAt` timestamp NOT NULL DEFAULT (now()) ON UPDATE CURRENT_TIMESTAMP,
	CONSTRAINT `drafts_id` PRIMARY KEY(`id`)
);
--> statement-breakpoint
CREATE TABLE `photos` (
	`id` int AUTO_INCREMENT NOT NULL,
	`reportId` int NOT NULL,
	`photoUrl` text NOT NULL,
	`photoKey` text NOT NULL,
	`caption` text,
	`order` int DEFAULT 0,
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	CONSTRAINT `photos_id` PRIMARY KEY(`id`)
);
--> statement-breakpoint
CREATE TABLE `reports` (
	`id` int AUTO_INCREMENT NOT NULL,
	`userId` int NOT NULL,
	`projectistName` text,
	`projectistCompany` text,
	`projectistPhone` text,
	`clientCompanyName` text,
	`clientAddress` text,
	`clientCity` text,
	`clientState` text,
	`clientZipCode` text,
	`clientNeighborhood` text,
	`clientSANumber` text,
	`clientARTNumber` text,
	`unitLocation` text,
	`isNewClient` enum('yes','no') DEFAULT 'yes',
	`isAsBuilt` enum('yes','no') DEFAULT 'no',
	`isAdequation` enum('yes','no') DEFAULT 'no',
	`checklist` json,
	`quantitatives` json,
	`logisticDifficulty` int,
	`assemblyDifficulty` int,
	`observations` text,
	`clientActions` text,
	`supergasActions` text,
	`status` enum('draft','completed') DEFAULT 'draft',
	`createdAt` timestamp NOT NULL DEFAULT (now()),
	`updatedAt` timestamp NOT NULL DEFAULT (now()) ON UPDATE CURRENT_TIMESTAMP,
	`completedAt` timestamp,
	CONSTRAINT `reports_id` PRIMARY KEY(`id`)
);
