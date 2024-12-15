import * as React from "react";
import { makeStyles } from "@fluentui/react-components";

interface SplashScreenProps {
  opacity: number;
}

const useStyles = makeStyles({
  splashScreen: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    height: "100vh",
    width: "100%",
    position: "fixed",
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
    backgroundImage: "url('../../assets/SplashScreen_background.png')",
    backgroundSize: "cover",
    backgroundPosition: "center",
    zIndex: 1000,
    transition: "opacity 0.5s ease-in-out",
  },
  spinner: {
    width: "10%",
    maxWidth: "40px",
    aspectRatio: "1 / 1",
    animation: "spin 3s linear infinite",
    marginTop: "15vh",
  },
});

const SplashScreen: React.FC<SplashScreenProps> = ({ opacity }) => {
  const styles = useStyles();

  React.useEffect(() => {
    const style = document.createElement("style");
    style.textContent = `
      @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
      }
    `;
    document.head.append(style);

    return () => {
      style.remove();
    };
  }, []);

  return (
    <div className={styles.splashScreen} style={{ opacity }}>
      <img src="../../assets/loading_spinner.png" alt="Loading" className={styles.spinner} />
    </div>
  );
};

export default SplashScreen;
