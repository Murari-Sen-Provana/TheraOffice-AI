import logging

def setup_logger(log_file='rpa_run.log'):
    """Configures and returns a logger that writes to a file and the console."""
    logger = logging.getLogger("TheraOfficeRPA")
    logger.setLevel(logging.DEBUG)  # DEBUG instead of INFO
    
    if logger.hasHandlers():
        logger.handlers.clear()
        
    fh = logging.FileHandler(log_file, mode='w')
    fh.setLevel(logging.DEBUG)
    
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)
    
    logger.addHandler(fh)
    logger.addHandler(ch)
    
    return logger
