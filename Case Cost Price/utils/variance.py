def calculate_variance(prediction, percent=5):
    """
    Returns +/- variance around predicted value.
    Default variance = ±5%
    """
    low = prediction * (1 - percent / 100)
    high = prediction * (1 + percent / 100)
    return low, high
