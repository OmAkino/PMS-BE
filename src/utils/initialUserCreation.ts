import { UserModel } from "../models/user.model";
import logger from "./logger";
import bcrypt from "bcryptjs";

const initialUserCreation = async () => {
  try {
    const superAdminEmail = "janu@gmail.com";
    const superAdminPassword = "Password@1";
    const superAdminFirstName = "Super";
    const superAdminLastName = "Admin";
    const superAdminRole = "superAdmin";

    // Check if super admin already exists
    const existingAdmin = await UserModel.findOne({ email: superAdminEmail });

    if (existingAdmin) {
      logger.info("Super Admin already exists. Skipping creation.");
      return;
    }

    // Hash password
    const hashedPassword = await bcrypt.hash(superAdminPassword, 10);

    // Create new super admin
    const newAdmin = new UserModel({
      firstName: superAdminFirstName,
      lastName: superAdminLastName,
      email: superAdminEmail,
      password: hashedPassword,
      role: superAdminRole,
    });

    await newAdmin.save();
    logger.info("Super Admin account created successfully!");
  } catch (err) {
    logger.error("Failed to create Super Admin: " + err);
  }
};
export default initialUserCreation;