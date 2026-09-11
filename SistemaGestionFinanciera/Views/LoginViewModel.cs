using System.ComponentModel.DataAnnotations;

namespace SistemaGestionFinanciera.ViewModels
{
    public class LoginViewModel
    {
        [Required(ErrorMessage = "Debe ingresar el correo.")]
        [EmailAddress(ErrorMessage = "Debe ingresar un correo válido.")]
        public string Email { get; set; } = string.Empty;

        [Required(ErrorMessage = "Debe ingresar la contraseña.")]
        public string Password { get; set; } = string.Empty;
    }
}