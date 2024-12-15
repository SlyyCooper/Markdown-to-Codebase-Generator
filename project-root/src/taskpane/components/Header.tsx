import * as React from "react";
import { Image, Text, Button, makeStyles } from "@fluentui/react-components";
import { Settings24Regular } from "@fluentui/react-icons";

interface HeaderProps {
  logo: string;
  title: string;
  onSettingsClick: () => void;
}

const useStyles = makeStyles({
  header: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "10px 20px",
    backgroundColor: "rgba(255, 255, 255, 0.7)",
    backdropFilter: "blur(10px)",
    borderBottom: "none",
    boxShadow: "0 1px 3px rgba(0,0,0,0.1)",
  },
  logoContainer: {
    display: "flex",
    alignItems: "center",
  },
  logo: {
    width: "40px",
    height: "40px",
    marginRight: "10px",
  },
  title: {
    fontSize: "18px",
    fontWeight: "bold",
  },
});

const Header: React.FC<HeaderProps> = ({ logo, title, onSettingsClick }) => {
  const styles = useStyles();

  return (
    <header className={styles.header}>
      <div className={styles.logoContainer}>
        <Image className={styles.logo} src={logo} alt="Logo" />
        <Text className={styles.title}>{title}</Text>
      </div>
      <Button icon={<Settings24Regular />} onClick={onSettingsClick} />
    </header>
  );
};

export default Header;
