using Teams_Backend_API.Models;
using Teams_Backend_API.Models.DTOs;

public interface IBookingService
{
    Task<Booking> CreateBookingAsync(BookingDto dto, string userEmail, string userName);
    Task<IReadOnlyList<Booking>> GetBookingsForUserAsync(string userEmail);
    Task<IReadOnlyList<Booking>> GetAllBookingsAsync();
}